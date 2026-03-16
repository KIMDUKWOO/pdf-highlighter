import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { PDFDocument, rgb } from "pdf-lib";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

const pdfInput = document.getElementById("pdfFile");
const excelInput = document.getElementById("excelFile");
const runBtn = document.getElementById("runBtn");
const logEl = document.getElementById("log");

const ignoreCaseEl = document.getElementById("ignoreCase");
const wholeWordEl = document.getElementById("wholeWord");
const bColorEl = document.getElementById("bColor");

let pdfFile = null;
let excelFile = null;

function log(msg) {
  logEl.textContent += msg + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}

function clearLog() {
  logEl.textContent = "";
}

function getColor(name) {
  const map = {
    green: rgb(0.6, 1.0, 0.6),
    blue: rgb(0.6, 0.85, 1.0),
    pink: rgb(1.0, 0.6, 0.85),
    orange: rgb(1.0, 0.75, 0.4),
    purple: rgb(0.75, 0.6, 1.0),
  };
  return map[name] || map.green;
}

function cloneBytes(srcUint8) {
  return srcUint8.slice();
}

function normalizeChar(ch) {
  if (!ch) return "";
  const map = {
    "‘": "'",
    "’": "'",
    "‚": "'",
    "‛": "'",
    "′": "'",
    "＇": "'",
    "`": "'",
    "´": "'",
    "ʹ": "'",
    "ˈ": "'",
    "“": '"',
    "”": '"',
    "„": '"',
    "‟": '"',
    "″": '"',
    "＂": '"',
    "–": "-",
    "—": "-",
    "−": "-",
    "‒": "-",
    "﹣": "-",
    "－": "-",
  };
  if (map[ch]) return map[ch];
  if (/\s/.test(ch)) return " ";
  return ch;
}

function normalizeText(text, ignoreCase = true) {
  let out = "";
  for (const ch of String(text || "")) out += normalizeChar(ch);
  out = out.replace(/\s+/g, " ").trim();
  if (ignoreCase) out = out.toLowerCase();
  return out;
}

function isWordChar(ch) {
  return /[a-zA-Z0-9_]/.test(ch || "");
}

function isBoundary(text, start, end) {
  const prev = start > 0 ? text[start - 1] : "";
  const next = end < text.length ? text[end] : "";
  return !isWordChar(prev) && !isWordChar(next);
}

function getTextItemInfo(item, viewportHeight) {
  const [a, b, c, d, e, f] = item.transform;
  const width = Math.max(item.width || 0, 1);
  const height = Math.max(Math.abs(item.height || d || 8), 1);

  const angle = Math.atan2(b, a) * 180 / Math.PI;
  const normalizedAngle = ((angle % 360) + 360) % 360;

  const horizontal =
    normalizedAngle < 15 ||
    normalizedAngle > 345 ||
    (normalizedAngle > 165 && normalizedAngle < 195);

  const x = e;
  const yTop = f;

  return {
    str: item.str || "",
    x,
    yTop,
    width,
    height,
    pdfY: viewportHeight - yTop - height,
    angle: normalizedAngle,
    horizontal,
  };
}

function buildHorizontalLines(items, viewportHeight) {
  const parsed = items
    .filter((it) => typeof it.str === "string" && it.str.trim() !== "")
    .map((it) => getTextItemInfo(it, viewportHeight))
    .filter((it) => it.horizontal);

  parsed.sort((a, b) => {
    const dy = Math.abs(a.yTop - b.yTop);
    if (dy > 2) return a.yTop - b.yTop;
    return a.x - b.x;
  });

  const lines = [];

  for (const item of parsed) {
    let line = null;

    for (const existing of lines) {
      const tol = Math.max(2, existing.avgHeight * 0.45);
      if (Math.abs(existing.yTop - item.yTop) <= tol) {
        line = existing;
        break;
      }
    }

    if (!line) {
      line = {
        yTop: item.yTop,
        avgHeight: item.height,
        items: [],
      };
      lines.push(line);
    }

    line.items.push(item);
    line.avgHeight =
      (line.avgHeight * (line.items.length - 1) + item.height) / line.items.length;
  }

  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
  }

  return lines;
}

function rectFromItem(item) {
  return {
    x: item.x,
    y: item.pdfY,
    width: Math.max(item.width, 1),
    height: Math.max(item.height, 1),
  };
}

function drawRect(pdfPage, rect, color) {
  // 비정상적으로 긴 세로 박스 방지
  if (rect.height > rect.width * 6 && rect.height > 40) {
    return;
  }

  pdfPage.drawRectangle({
    x: rect.x,
    y: rect.y,
    width: rect.width,
    height: rect.height,
    color,
    opacity: 0.35,
    borderWidth: 0,
  });
}

function drawRects(pdfPage, rects, color) {
  for (const rect of rects) {
    drawRect(pdfPage, rect, color);
  }
}

function readKeywordsFromWorkbook(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const seenA = new Set();
  const seenB = new Set();
  const aKeywords = [];
  const bKeywords = [];

  for (const row of rows) {
    const a = row?.[0] != null ? String(row[0]).trim() : "";
    const b = row?.[1] != null ? String(row[1]).trim() : "";

    if (a && !seenA.has(a)) {
      seenA.add(a);
      aKeywords.push(a);
    }
    if (b && !seenB.has(b)) {
      seenB.add(b);
      bKeywords.push(b);
    }
  }

  return { aKeywords, bKeywords };
}

function findItemExactMatches(lineItems, keywordNorm, wholeWord, ignoreCase) {
  const hits = [];

  for (const item of lineItems) {
    const norm = normalizeText(item.str, ignoreCase);
    if (!norm) continue;

    if (norm === keywordNorm) {
      if (!wholeWord || isBoundary(` ${norm} `, 1, 1 + norm.length)) {
        hits.push([item]);
      }
    }
  }

  return hits;
}

function findMultiItemMatches(lineItems, keywordNorm, wholeWord, ignoreCase) {
  const hits = [];

  for (let start = 0; start < lineItems.length; start++) {
    let combined = "";
    const picked = [];

    for (let end = start; end < lineItems.length; end++) {
      const item = lineItems[end];
      const itemNorm = normalizeText(item.str, ignoreCase);
      if (!itemNorm) continue;

      if (picked.length > 0) {
        const prev = picked[picked.length - 1];
        const gap = item.x - (prev.x + prev.width);
        const gapLimit = Math.max(8, Math.min(prev.height, item.height) * 0.8);

        // 너무 멀면 같은 문자열로 보지 않음
        if (gap > gapLimit) break;

        // 하이픈 계열/공백 허용
        if (gap > 1.5) combined += " ";
      }

      combined += itemNorm;
      picked.push(item);

      const compactCombined = combined.replace(/\s+/g, "");
      const compactKeyword = keywordNorm.replace(/\s+/g, "");

      if (compactCombined.length > compactKeyword.length + 6) break;

      if (compactCombined === compactKeyword) {
        if (!wholeWord || isBoundary(` ${compactCombined} `, 1, 1 + compactCombined.length)) {
          hits.push([...picked]);
        }
        break;
      }
    }
  }

  return hits;
}

function removeDuplicateHits(hitGroups) {
  const seen = new Set();
  const result = [];

  for (const group of hitGroups) {
    const key = group.map((it) => `${it.x.toFixed(2)}|${it.yTop.toFixed(2)}|${it.str}`).join("||");
    if (!seen.has(key)) {
      seen.add(key);
      result.push(group);
    }
  }

  return result;
}

pdfInput.addEventListener("change", (e) => {
  pdfFile = e.target.files[0] || null;
  runBtn.disabled = !(pdfFile && excelFile);
});

excelInput.addEventListener("change", (e) => {
  excelFile = e.target.files[0] || null;
  runBtn.disabled = !(pdfFile && excelFile);
});

runBtn.addEventListener("click", async () => {
  clearLog();

  try {
    if (!pdfFile) throw new Error("PDF 파일이 선택되지 않았습니다.");
    if (!excelFile) throw new Error("엑셀 파일이 선택되지 않았습니다.");

    runBtn.disabled = true;

    log("엑셀 분석 중...");

    const excelData = await excelFile.arrayBuffer();
    const workbook = XLSX.read(excelData, { type: "array" });
    const { aKeywords, bKeywords } = readKeywordsFromWorkbook(workbook);

    log(`A키워드 ${aKeywords.length}개`);
    log(`B키워드 ${bKeywords.length}개`);

    if (aKeywords.length === 0 && bKeywords.length === 0) {
      throw new Error("엑셀 A열/B열에서 키워드를 찾지 못했습니다.");
    }

    log("PDF 로딩 중...");

    const originalPdfBuffer = await pdfFile.arrayBuffer();
    const originalPdfBytes = new Uint8Array(originalPdfBuffer);
    const pdfBytesForPdfJs = cloneBytes(originalPdfBytes);
    const pdfBytesForPdfLib = cloneBytes(originalPdfBytes);

    const pdfJsDoc = await pdfjsLib.getDocument({ data: pdfBytesForPdfJs }).promise;
    const pdfLibDoc = await PDFDocument.load(pdfBytesForPdfLib);

    const ignoreCase = ignoreCaseEl.checked;
    const wholeWord = wholeWordEl.checked;
    const aColor = rgb(1, 1, 0);
    const bColor = getColor(bColorEl.value);

    const groups = [
      {
        label: "A",
        color: aColor,
        keywords: aKeywords.map((raw) => ({ raw, norm: normalizeText(raw, ignoreCase) })).filter((k) => k.norm),
      },
      {
        label: "B",
        color: bColor,
        keywords: bKeywords.map((raw) => ({ raw, norm: normalizeText(raw, ignoreCase) })).filter((k) => k.norm),
      },
    ];

    let totalMatches = 0;

    for (let pageIndex = 0; pageIndex < pdfJsDoc.numPages; pageIndex++) {
      log(`페이지 ${pageIndex + 1}/${pdfJsDoc.numPages} 처리 중...`);

      const page = await pdfJsDoc.getPage(pageIndex + 1);
      const viewport = page.getViewport({ scale: 1 });
      const textContent = await page.getTextContent();
      const lines = buildHorizontalLines(textContent.items || [], viewport.height);
      const pdfPage = pdfLibDoc.getPage(pageIndex);

      for (const group of groups) {
        for (const kw of group.keywords) {
          const allHits = [];

          for (const line of lines) {
            const exactHits = findItemExactMatches(line.items, kw.norm, wholeWord, ignoreCase);
            const multiHits = exactHits.length
              ? []
              : findMultiItemMatches(line.items, kw.norm, wholeWord, ignoreCase);

            allHits.push(...exactHits, ...multiHits);
          }

          const uniqueHits = removeDuplicateHits(allHits);

          for (const hit of uniqueHits) {
            const rects = hit.map(rectFromItem);
            drawRects(pdfPage, rects, group.color);
            totalMatches += 1;
          }
        }
      }
    }

    log(`총 매칭 수: ${totalMatches}`);
    log("PDF 저장 중...");

    const resultBytes = await pdfLibDoc.save();
    const blob = new Blob([resultBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = pdfFile.name.replace(/\.pdf$/i, "") + "_highlighted.pdf";
    a.click();

    setTimeout(() => URL.revokeObjectURL(url), 1000);

    log("완료!");
  } catch (err) {
    console.error(err);
    log("");
    log("[오류]");
    log(err?.message || String(err));
    alert(`오류가 발생했습니다.\n${err?.message || String(err)}`);
  } finally {
    runBtn.disabled = !(pdfFile && excelFile);
  }
});
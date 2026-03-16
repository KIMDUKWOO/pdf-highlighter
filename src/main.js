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

  const quoteMap = {
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
    "˝": '"',
    "“": '"',
    "”": '"',
    "„": '"',
    "‟": '"',
    "″": '"',
    "＂": '"',
    "–": "-",
    "—": "-",
    "−": "-",
    "-": "-",
    "‒": "-",
    "﹣": "-",
    "－": "-",
  };

  if (quoteMap[ch]) return quoteMap[ch];
  if (/\s/.test(ch)) return " ";

  return ch;
}

function normalizeText(text, ignoreCase = true) {
  let out = "";
  for (const ch of String(text || "")) {
    out += normalizeChar(ch);
  }

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

function getTextItemBox(item, viewportHeight) {
  const t = item.transform;
  const x = t[4];
  const yTop = t[5];
  const width = item.width || 0;
  const height = Math.abs(item.height || t[0] || 10);

  return {
    x,
    yTop,
    width,
    height,
    // pdf-lib 좌표계(bottom-left origin)로 변환
    pdfY: viewportHeight - yTop - height,
  };
}

function groupItemsIntoLines(items, viewportHeight) {
  const usable = items
    .filter((it) => typeof it.str === "string" && it.str.length > 0)
    .map((it, idx) => {
      const box = getTextItemBox(it, viewportHeight);
      return {
        rawIndex: idx,
        str: it.str,
        x: box.x,
        yTop: box.yTop,
        width: box.width,
        height: box.height,
        pdfY: box.pdfY,
      };
    });

  usable.sort((a, b) => {
    const dy = Math.abs(a.yTop - b.yTop);
    if (dy > 2) return a.yTop - b.yTop;
    return a.x - b.x;
  });

  const lines = [];
  for (const item of usable) {
    let found = null;
    for (const line of lines) {
      const tolerance = Math.max(2, line.avgHeight * 0.5);
      if (Math.abs(line.yTop - item.yTop) <= tolerance) {
        found = line;
        break;
      }
    }

    if (!found) {
      found = {
        yTop: item.yTop,
        avgHeight: item.height || 10,
        items: [],
      };
      lines.push(found);
    }

    found.items.push(item);
    found.avgHeight =
      (found.avgHeight * (found.items.length - 1) + (item.height || 10)) /
      found.items.length;
  }

  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
  }

  return lines;
}

function buildSearchableLine(lineItems, ignoreCase) {
  let text = "";
  const charMap = [];

  for (let i = 0; i < lineItems.length; i++) {
    const item = lineItems[i];
    const prev = i > 0 ? lineItems[i - 1] : null;

    if (prev) {
      const prevRight = prev.x + prev.width;
      const gap = item.x - prevRight;
      const threshold = Math.max(1.5, Math.min(prev.height, item.height) * 0.15);

      if (gap > threshold) {
        text += " ";
        charMap.push(-1);
      }
    }

    const normalized = normalizeText(item.str, ignoreCase);

    for (const ch of normalized) {
      text += ch;
      charMap.push(i);
    }
  }

  return { text, charMap };
}

function findAllMatches(searchText, keyword, wholeWord) {
  const matches = [];
  if (!keyword || !searchText) return matches;

  let fromIndex = 0;
  while (fromIndex < searchText.length) {
    const hit = searchText.indexOf(keyword, fromIndex);
    if (hit === -1) break;

    const end = hit + keyword.length;
    if (!wholeWord || isBoundary(searchText, hit, end)) {
      matches.push({ start: hit, end });
    }

    fromIndex = hit + 1;
  }

  return matches;
}

function rectFromItem(item) {
  return {
    x: item.x,
    y: item.pdfY,
    width: Math.max(item.width, 1),
    height: Math.max(item.height, 1),
  };
}

function mergeRects(rects) {
  if (!rects.length) return [];

  const sorted = [...rects].sort((a, b) => {
    const dy = Math.abs(a.y - b.y);
    if (dy > 2) return b.y - a.y;
    return a.x - b.x;
  });

  const merged = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    const last = merged[merged.length - 1];

    const sameLine = Math.abs(cur.y - last.y) <= Math.max(2, Math.min(cur.height, last.height) * 0.5);
    const close = cur.x <= last.x + last.width + 3;

    if (sameLine && close) {
      const newRight = Math.max(last.x + last.width, cur.x + cur.width);
      const newTop = Math.max(last.y + last.height, cur.y + cur.height);
      last.x = Math.min(last.x, cur.x);
      last.y = Math.min(last.y, cur.y);
      last.width = newRight - last.x;
      last.height = newTop - last.y;
    } else {
      merged.push({ ...cur });
    }
  }

  return merged;
}

function getRectsForMatch(lineItems, charMap, match) {
  const itemIndexes = new Set();

  for (let i = match.start; i < match.end; i++) {
    const idx = charMap[i];
    if (idx >= 0) itemIndexes.add(idx);
  }

  const rects = [...itemIndexes]
    .sort((a, b) => a - b)
    .map((idx) => rectFromItem(lineItems[idx]));

  return mergeRects(rects);
}

function drawHighlightRects(pdfPage, rects, color) {
  for (const r of rects) {
    pdfPage.drawRectangle({
      x: r.x,
      y: r.y,
      width: r.width,
      height: r.height,
      color,
      opacity: 0.35,
      borderWidth: 0,
    });
  }
}

function readKeywordsFromWorkbook(workbook) {
  if (!workbook.SheetNames.length) {
    throw new Error("엑셀 시트를 찾을 수 없습니다.");
  }

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

    const pdfJsDoc = await pdfjsLib.getDocument({
      data: pdfBytesForPdfJs,
    }).promise;

    const pdfLibDoc = await PDFDocument.load(pdfBytesForPdfLib);

    log(`PDF 페이지 수: ${pdfJsDoc.numPages}`);
    log("정확 좌표 검색 시작...");

    const ignoreCase = ignoreCaseEl.checked;
    const wholeWord = wholeWordEl.checked;
    const bColor = getColor(bColorEl.value);
    const aColor = rgb(1, 1, 0);

    const normalizedA = aKeywords.map((k) => ({
      raw: k,
      norm: normalizeText(k, ignoreCase),
    })).filter((x) => x.norm);

    const normalizedB = bKeywords.map((k) => ({
      raw: k,
      norm: normalizeText(k, ignoreCase),
    })).filter((x) => x.norm);

    let totalMatches = 0;
    const keywordStats = new Map();

    for (const k of normalizedA) keywordStats.set(`A:${k.raw}`, 0);
    for (const k of normalizedB) keywordStats.set(`B:${k.raw}`, 0);

    for (let pageIndex = 0; pageIndex < pdfJsDoc.numPages; pageIndex++) {
      log(`페이지 ${pageIndex + 1}/${pdfJsDoc.numPages} 처리 중...`);

      const page = await pdfJsDoc.getPage(pageIndex + 1);
      const viewport = page.getViewport({ scale: 1.0 });
      const textContent = await page.getTextContent();
      const pdfPage = pdfLibDoc.getPage(pageIndex);

      const lines = groupItemsIntoLines(textContent.items || [], viewport.height);

      for (const line of lines) {
        const { text, charMap } = buildSearchableLine(line.items, ignoreCase);
        if (!text) continue;

        for (const kw of normalizedA) {
          const matches = findAllMatches(text, kw.norm, wholeWord);
          if (!matches.length) continue;

          for (const match of matches) {
            const rects = getRectsForMatch(line.items, charMap, match);
            drawHighlightRects(pdfPage, rects, aColor);
            totalMatches += 1;
            keywordStats.set(`A:${kw.raw}`, (keywordStats.get(`A:${kw.raw}`) || 0) + 1);
          }
        }

        for (const kw of normalizedB) {
          const matches = findAllMatches(text, kw.norm, wholeWord);
          if (!matches.length) continue;

          for (const match of matches) {
            const rects = getRectsForMatch(line.items, charMap, match);
            drawHighlightRects(pdfPage, rects, bColor);
            totalMatches += 1;
            keywordStats.set(`B:${kw.raw}`, (keywordStats.get(`B:${kw.raw}`) || 0) + 1);
          }
        }
      }
    }

    const topHits = [...keywordStats.entries()]
      .filter(([, count]) => count > 0)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 20);

    log(`총 매칭 수: ${totalMatches}`);
    if (topHits.length) {
      log("상위 매칭:");
      for (const [name, count] of topHits) {
        log(`- ${name} : ${count}`);
      }
    } else {
      log("매칭된 키워드가 없습니다.");
    }

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
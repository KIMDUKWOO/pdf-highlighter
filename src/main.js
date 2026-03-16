import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import workerSrc from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { PDFDocument, rgb } from "pdf-lib";

pdfjsLib.GlobalWorkerOptions.workerSrc = workerSrc;

const pdfInput = document.getElementById("pdfFile");
const excelInput = document.getElementById("excelFile");
const runBtn = document.getElementById("runBtn");
const cancelBtn = document.getElementById("cancelBtn");
const logEl = document.getElementById("log");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");

const ignoreCaseEl = document.getElementById("ignoreCase");
const wholeWordEl = document.getElementById("wholeWord");
const bColorEl = document.getElementById("bColor");

let pdfFile = null;
let excelFile = null;
let shouldCancel = false;

function log(msg) {
  logEl.textContent += msg + "\n";
  logEl.scrollTop = logEl.scrollHeight;
}

function clearLog() {
  logEl.textContent = "";
}

function setRunning(running) {
  runBtn.disabled = running || !(pdfFile && excelFile);
  cancelBtn.disabled = !running;
}

function setProgress(done, total, text) {
  const pct = total > 0 ? (done / total) * 100 : 0;
  progressBar.value = pct;
  progressText.textContent = `${pct.toFixed(1)}% | ${text}`;
}

function getColor(name) {
  const map = {
    green: rgb(0.55, 1.0, 0.55),
    blue: rgb(0.55, 0.85, 1.0),
    pink: rgb(1.0, 0.60, 0.85),
    orange: rgb(1.0, 0.75, 0.40),
    purple: rgb(0.75, 0.60, 1.0),
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

function compactText(text, ignoreCase = true) {
  return normalizeText(text, ignoreCase).replace(/\s+/g, "");
}

function isWordChar(ch) {
  return /[A-Za-z0-9_]/.test(ch || "");
}

function isBoundary(text, start, end) {
  const prev = start > 0 ? text[start - 1] : "";
  const next = end < text.length ? text[end] : "";
  return !isWordChar(prev) && !isWordChar(next);
}

function getItemInfo(item, viewportHeight) {
  const [a, b, c, d, e, f] = item.transform;
  const width = Math.max(item.width || 1, 1);
  const height = Math.max(Math.abs(item.height || d || 8), 1);

  const angle = (Math.atan2(b, a) * 180) / Math.PI;
  const normalizedAngle = ((angle % 360) + 360) % 360;

  const horizontal =
    normalizedAngle < 15 ||
    normalizedAngle > 345 ||
    (normalizedAngle > 165 && normalizedAngle < 195);

  const x = e;
  const yTop = f;
  const pdfY = viewportHeight - yTop - height;

  return {
    str: item.str || "",
    x,
    yTop,
    width,
    height,
    pdfY,
    angle: normalizedAngle,
    horizontal,
  };
}

function groupItemsIntoLines(items, viewportHeight) {
  const parsed = items
    .filter((it) => typeof it.str === "string" && it.str.length > 0)
    .map((it) => getItemInfo(it, viewportHeight))
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
      const tol = Math.max(2, existing.avgHeight * 0.5);
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

function buildCharLevelLine(lineItems, ignoreCase) {
  let text = "";
  const charBoxes = [];

  for (let i = 0; i < lineItems.length; i++) {
    const item = lineItems[i];
    const prev = i > 0 ? lineItems[i - 1] : null;

    if (prev) {
      const prevRight = prev.x + prev.width;
      const gap = item.x - prevRight;
      const threshold = Math.max(1.5, Math.min(prev.height, item.height) * 0.15);

      if (gap > threshold) {
        text += " ";
        charBoxes.push(null);
      }
    }

    const raw = String(item.str || "");
    const normalizedChars = [...raw].map((ch) => normalizeChar(ch));
    const itemText = normalizedChars.join("");
    const itemNorm = ignoreCase ? itemText.toLowerCase() : itemText;

    const glyphCount = Math.max(itemNorm.length, 1);
    const charWidth = item.width / glyphCount;

    for (let idx = 0; idx < itemNorm.length; idx++) {
      const ch = itemNorm[idx];
      text += ch;

      charBoxes.push({
        x: item.x + charWidth * idx,
        y: item.pdfY,
        width: Math.max(charWidth, 1),
        height: item.height,
      });
    }
  }

  text = text.replace(/\s+/g, " ");

  return { text, charBoxes };
}

function splitLineToWordTokens(lineItems, ignoreCase) {
  const tokens = [];

  for (const item of lineItems) {
    const raw = String(item.str || "");
    const norm = normalizeText(raw, ignoreCase);
    if (!norm) continue;

    const parts = norm.split(" ").filter(Boolean);
    if (!parts.length) continue;

    const tokenWidth = item.width / parts.length;

    for (let i = 0; i < parts.length; i++) {
      tokens.push({
        text: parts[i],
        x: item.x + tokenWidth * i,
        y: item.pdfY,
        width: Math.max(tokenWidth, 1),
        height: item.height,
      });
    }
  }

  return tokens;
}

function rectUnion(rects) {
  if (!rects.length) return null;

  let x0 = rects[0].x;
  let y0 = rects[0].y;
  let x1 = rects[0].x + rects[0].width;
  let y1 = rects[0].y + rects[0].height;

  for (let i = 1; i < rects.length; i++) {
    const r = rects[i];
    x0 = Math.min(x0, r.x);
    y0 = Math.min(y0, r.y);
    x1 = Math.max(x1, r.x + r.width);
    y1 = Math.max(y1, r.y + r.height);
  }

  return {
    x: x0,
    y: y0,
    width: x1 - x0,
    height: y1 - y0,
  };
}

function mergeNeighborRects(rects) {
  if (!rects.length) return [];

  const sorted = [...rects].sort((a, b) => {
    const dy = Math.abs(a.y - b.y);
    if (dy > 2) return b.y - a.y;
    return a.x - b.x;
  });

  const merged = [{ ...sorted[0] }];

  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    const last = merged[merged.length - 1];

    const sameLine = Math.abs(cur.y - last.y) <= Math.max(2, Math.min(cur.height, last.height) * 0.5);
    const close = cur.x <= last.x + last.width + 2;

    if (sameLine && close) {
      const right = Math.max(last.x + last.width, cur.x + cur.width);
      const top = Math.max(last.y + last.height, cur.y + cur.height);
      last.x = Math.min(last.x, cur.x);
      last.y = Math.min(last.y, cur.y);
      last.width = right - last.x;
      last.height = top - last.y;
    } else {
      merged.push({ ...cur });
    }
  }

  return merged;
}

function drawHighlightRects(pdfPage, rects, color) {
  for (const r of rects) {
    if (!r) continue;

    if (r.height > r.width * 6 && r.height > 40) {
      continue;
    }

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

  const aKeywords = [];
  const bKeywords = [];
  const seenA = new Set();
  const seenB = new Set();

  for (const row of rows) {
    const aVal = row?.[0];
    const bVal = row?.[1];

    if (aVal != null) {
      const kw = String(aVal).trim();
      if (kw && !seenA.has(kw)) {
        seenA.add(kw);
        aKeywords.push(kw);
      }
    }

    if (bVal != null) {
      const kw = String(bVal).trim();
      if (kw && !seenB.has(kw)) {
        seenB.add(kw);
        bKeywords.push(kw);
      }
    }
  }

  return { aKeywords, bKeywords };
}

function findSubstringMatches(charLineText, keywordNorm, wholeWord) {
  const matches = [];
  if (!keywordNorm || !charLineText) return matches;

  let startIndex = 0;
  while (startIndex < charLineText.length) {
    const hit = charLineText.indexOf(keywordNorm, startIndex);
    if (hit === -1) break;

    const end = hit + keywordNorm.length;
    if (!wholeWord || isBoundary(charLineText, hit, end)) {
      matches.push({ start: hit, end });
    }

    startIndex = hit + 1;
  }

  return matches;
}

function rectsFromCharMatch(charBoxes, match) {
  const rects = [];

  for (let i = match.start; i < match.end; i++) {
    const box = charBoxes[i];
    if (box) rects.push(box);
  }

  return mergeNeighborRects(rects);
}

function findWordPhraseMatches(tokens, keywordRaw, ignoreCase) {
  const parts = normalizeText(keywordRaw, ignoreCase).split(" ").filter(Boolean);
  if (!parts.length) return [];

  const hits = [];

  for (let i = 0; i <= tokens.length - parts.length; i++) {
    let ok = true;
    for (let j = 0; j < parts.length; j++) {
      if (tokens[i + j].text !== parts[j]) {
        ok = false;
        break;
      }
    }

    if (ok) {
      hits.push(tokens.slice(i, i + parts.length));
    }
  }

  return hits;
}

function makeStats(aKeywords, bKeywords) {
  const perKeyword = {};
  for (const k of aKeywords) perKeyword[k] = 0;
  for (const k of bKeywords) {
    if (k in perKeyword) perKeyword[`(B)${k}`] = 0;
    else perKeyword[k] = 0;
  }
  return perKeyword;
}

pdfInput.addEventListener("change", (e) => {
  pdfFile = e.target.files?.[0] || null;
  runBtn.disabled = !(pdfFile && excelFile);
});

excelInput.addEventListener("change", (e) => {
  excelFile = e.target.files?.[0] || null;
  runBtn.disabled = !(pdfFile && excelFile);
});

cancelBtn.addEventListener("click", () => {
  shouldCancel = true;
  progressText.textContent = "취소 요청 중...";
});

runBtn.addEventListener("click", async () => {
  clearLog();
  shouldCancel = false;

  try {
    if (!pdfFile) throw new Error("PDF 파일을 선택해 주세요.");
    if (!excelFile) throw new Error("엑셀 파일을 선택해 주세요.");

    setRunning(true);
    setProgress(0, 1, "엑셀 키워드 로딩 중...");

    log("엑셀 분석 중...");

    const excelData = await excelFile.arrayBuffer();
    const workbook = XLSX.read(excelData, { type: "array" });
    const { aKeywords, bKeywords } = readKeywordsFromWorkbook(workbook);

    if (!aKeywords.length && !bKeywords.length) {
      throw new Error("엑셀 A열/B열에서 키워드를 찾지 못했습니다.");
    }

    log(`A키워드 ${aKeywords.length}개`);
    log(`B키워드 ${bKeywords.length}개`);

    setProgress(0, 1, "PDF 분석 준비...");

    const originalPdfBuffer = await pdfFile.arrayBuffer();
    const originalPdfBytes = new Uint8Array(originalPdfBuffer);

    const pdfBytesForPdfJs = cloneBytes(originalPdfBytes);
    const pdfBytesForPdfLib = cloneBytes(originalPdfBytes);

    const pdfJsDoc = await pdfjsLib.getDocument({
      data: pdfBytesForPdfJs,
    }).promise;

    const pdfLibDoc = await PDFDocument.load(pdfBytesForPdfLib);

    const numPages = pdfJsDoc.numPages;
    const allKeywordCount = Math.max(1, aKeywords.length + bKeywords.length);
    const totalSteps = numPages * allKeywordCount;

    const ignoreCase = ignoreCaseEl.checked;
    const wholeWord = wholeWordEl.checked;
    const aColor = rgb(1, 1, 0);
    const bColor = getColor(bColorEl.value);

    const perKeyword = makeStats(aKeywords, bKeywords);
    let totalMatches = 0;
    let doneSteps = 0;

    function report(message) {
      setProgress(doneSteps, totalSteps, message);
    }

    report("처리 시작...");

    const updateEvery = Math.max(1, Math.floor(totalSteps / 300));

    for (let pageIndex = 0; pageIndex < numPages; pageIndex++) {
      if (shouldCancel) {
        throw new Error("사용자가 작업을 취소했습니다.");
      }

      const page = await pdfJsDoc.getPage(pageIndex + 1);
      const viewport = page.getViewport({ scale: 1.0 });
      const textContent = await page.getTextContent();
      const lines = groupItemsIntoLines(textContent.items || [], viewport.height);
      const pdfPage = pdfLibDoc.getPage(pageIndex);

      async function processKeywordList(keywords, color, labelPrefix = "") {
        for (const kw of keywords) {
          if (shouldCancel) {
            throw new Error("사용자가 작업을 취소했습니다.");
          }

          let hitCount = 0;

          if (!wholeWord) {
            const keywordNorm = compactText(kw, ignoreCase);

            for (const line of lines) {
              const { text, charBoxes } = buildCharLevelLine(line.items, ignoreCase);
              const compactLine = text.replace(/\s+/g, "");

              if (!compactLine) continue;

              const compactMatches = findSubstringMatches(compactLine, keywordNorm, false);
              if (!compactMatches.length) continue;

              const compactToOriginal = [];
              for (let i = 0; i < text.length; i++) {
                if (text[i] !== " ") compactToOriginal.push(i);
              }

              for (const cm of compactMatches) {
                const originalStart = compactToOriginal[cm.start];
                const originalEnd = compactToOriginal[cm.end - 1] + 1;

                const rects = rectsFromCharMatch(charBoxes, {
                  start: originalStart,
                  end: originalEnd,
                });

                if (rects.length) {
                  drawHighlightRects(pdfPage, rects, color);
                  hitCount += 1;
                }
              }
            }
          } else {
            const kwParts = normalizeText(kw, ignoreCase).split(" ").filter(Boolean);

            for (const line of lines) {
              const tokens = splitLineToWordTokens(line.items, ignoreCase);
              if (!tokens.length) continue;

              if (kwParts.length === 1) {
                for (const token of tokens) {
                  if (token.text === kwParts[0]) {
                    drawHighlightRects(pdfPage, [token], color);
                    hitCount += 1;
                  }
                }
              } else {
                const hits = findWordPhraseMatches(tokens, kw, ignoreCase);
                for (const group of hits) {
                  const unionRect = rectUnion(group);
                  if (unionRect) {
                    drawHighlightRects(pdfPage, [unionRect], color);
                    hitCount += 1;
                  }
                }
              }
            }
          }

          if (hitCount > 0) {
            const keyName = labelPrefix ? `${labelPrefix}${kw}` : kw;
            perKeyword[keyName] = (perKeyword[keyName] || 0) + hitCount;
            totalMatches += hitCount;
          }

          doneSteps += 1;
          if (doneSteps % updateEvery === 0 || doneSteps === totalSteps) {
            report(`${pageIndex + 1}/${numPages} 페이지 처리 중... (키워드: ${kw})`);
          }

          await new Promise((resolve) => setTimeout(resolve, 0));
        }
      }

      await processKeywordList(aKeywords, aColor, "");
      await processKeywordList(bKeywords, bColor, "(B)");
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

    const stats = {
      total_matches: totalMatches,
      per_keyword: perKeyword,
      total_steps: totalSteps,
      pages: numPages,
      a_keywords: aKeywords.length,
      b_keywords: bKeywords.length,
    };

    const topHits = Object.entries(stats.per_keyword)
      .sort((x, y) => y[1] - x[1])
      .slice(0, 10);

    const topLines =
      topHits.filter(([, v]) => v > 0).map(([k, v]) => `- ${k}: ${v}회`).join("\n") ||
      "(매칭된 키워드 없음)";

    setProgress(totalSteps, totalSteps, "완료!");
    log(`총 하이라이트 개수: ${stats.total_matches}`);
    log(topLines);

    alert(
      `저장 완료\n\n총 하이라이트 개수: ${stats.total_matches}\n페이지: ${stats.pages} | A키워드: ${stats.a_keywords} | B키워드: ${stats.b_keywords}\n\n상위 매칭(최대 10개):\n${topLines}`
    );
  } catch (err) {
    console.error(err);
    log("");
    log("[오류]");
    log(err?.message || String(err));
    alert(`오류가 발생했습니다.\n${err?.message || String(err)}`);
    progressText.textContent = "오류 발생";
  } finally {
    setRunning(false);
  }
});
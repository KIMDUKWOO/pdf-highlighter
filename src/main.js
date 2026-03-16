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

function escapeRegExp(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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

function cloneUint8Array(src) {
  return new Uint8Array(src);
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

    log(`A키워드 ${aKeywords.length}개`);
    log(`B키워드 ${bKeywords.length}개`);

    if (aKeywords.length === 0 && bKeywords.length === 0) {
      throw new Error("엑셀 A열/B열에서 키워드를 찾지 못했습니다.");
    }

    log("PDF 로딩 중...");

    // 같은 ArrayBuffer를 재사용하지 않도록 복사본 2개 생성
    const originalPdfBuffer = await pdfFile.arrayBuffer();
    const pdfBytesForPdfJs = cloneUint8Array(originalPdfBuffer);
    const pdfBytesForPdfLib = cloneUint8Array(originalPdfBuffer);

    const loadingTask = pdfjsLib.getDocument({ data: pdfBytesForPdfJs });
    const pdfJsDoc = await loadingTask.promise;

    const pdfLibDoc = await PDFDocument.load(pdfBytesForPdfLib);

    log(`PDF 페이지 수: ${pdfJsDoc.numPages}`);
    log("텍스트 검색 시작...");

    const ignoreCase = ignoreCaseEl.checked;
    const wholeWord = wholeWordEl.checked;
    const bColor = getColor(bColorEl.value);

    let totalMatches = 0;

    for (let pageIndex = 0; pageIndex < pdfJsDoc.numPages; pageIndex++) {
      log(`페이지 ${pageIndex + 1}/${pdfJsDoc.numPages} 처리 중...`);

      const page = await pdfJsDoc.getPage(pageIndex + 1);
      const textContent = await page.getTextContent();
      const items = textContent.items || [];

      const pageText = items.map((item) => item.str || "").join(" ");
      const pdfPage = pdfLibDoc.getPage(pageIndex);
      const pageHeight = pdfPage.getHeight();

      const handleKeywordGroup = (keywords, color, label) => {
        for (const keyword of keywords) {
          const safeKeyword = escapeRegExp(keyword);
          const flags = ignoreCase ? "gi" : "g";
          const pattern = wholeWord
            ? new RegExp(`\\b${safeKeyword}\\b`, flags)
            : new RegExp(safeKeyword, flags);

          const matches = pageText.match(pattern);
          if (!matches || matches.length === 0) continue;

          totalMatches += matches.length;
          log(`  - ${label} "${keyword}" : ${matches.length}건`);

          // 현재는 임시 표시용 마킹
          const markerY = pageHeight - 40 - (totalMatches % 20) * 16;

          pdfPage.drawRectangle({
            x: 30,
            y: Math.max(20, markerY),
            width: 180,
            height: 12,
            color,
            opacity: 0.35,
          });
        }
      };

      handleKeywordGroup(aKeywords, rgb(1, 1, 0), "A열");
      handleKeywordGroup(bKeywords, bColor, "B열");
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
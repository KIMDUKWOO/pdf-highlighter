import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import { PDFDocument, rgb } from "pdf-lib";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

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

pdfInput.addEventListener("change", e => {
  pdfFile = e.target.files[0];
  if (pdfFile && excelFile) runBtn.disabled = false;
});

excelInput.addEventListener("change", e => {
  excelFile = e.target.files[0];
  if (pdfFile && excelFile) runBtn.disabled = false;
});

function getColor(name) {
  const map = {
    green: rgb(0.6, 1, 0.6),
    blue: rgb(0.6, 0.8, 1),
    pink: rgb(1, 0.6, 0.8),
    orange: rgb(1, 0.7, 0.4),
    purple: rgb(0.7, 0.6, 1),
  };
  return map[name];
}

runBtn.addEventListener("click", async () => {
  log("엑셀 분석 중...");

  const excelData = await excelFile.arrayBuffer();
  const workbook = XLSX.read(excelData);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const aKeywords = [];
  const bKeywords = [];

  json.forEach(row => {
    if (row[0]) aKeywords.push(String(row[0]).trim());
    if (row[1]) bKeywords.push(String(row[1]).trim());
  });

  log(`A키워드 ${aKeywords.length}개`);
  log(`B키워드 ${bKeywords.length}개`);

  const pdfBytes = await pdfFile.arrayBuffer();
  const pdfDoc = await PDFDocument.load(pdfBytes);
  const pdf = await pdfjsLib.getDocument({ data: pdfBytes }).promise;

  const ignoreCase = ignoreCaseEl.checked;
  const wholeWord = wholeWordEl.checked;
  const bColor = getColor(bColorEl.value);

  for (let i = 0; i < pdf.numPages; i++) {
    log(`페이지 ${i + 1} 처리 중...`);

    const page = await pdf.getPage(i + 1);
    const textContent = await page.getTextContent();
    const strings = textContent.items.map(item => item.str);
    const fullText = strings.join(" ");

    const pdfPage = pdfDoc.getPage(i);

    function highlight(keyword, color) {
      let flags = ignoreCase ? "gi" : "g";
      let pattern = wholeWord
        ? new RegExp(`\\b${keyword}\\b`, flags)
        : new RegExp(keyword, flags);

      let match;
      while ((match = pattern.exec(fullText)) !== null) {
        pdfPage.drawRectangle({
          x: 50,
          y: 50,
          width: 200,
          height: 15,
          color,
          opacity: 0.4,
        });
      }
    }

    aKeywords.forEach(k => highlight(k, rgb(1, 1, 0)));
    bKeywords.forEach(k => highlight(k, bColor));
  }

  const modifiedPdf = await pdfDoc.save();
  const blob = new Blob([modifiedPdf], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "highlighted.pdf";
  a.click();

  log("완료!");
});
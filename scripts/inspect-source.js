const XLSX = require("xlsx-js-style");
const Core = require("../bom-core.js");

const source = process.argv[2];
if (!source) throw new Error("Please provide an XLSX path.");
const workbook = XLSX.readFile(source, { cellStyles: true });
const analysis = Core.analyzeWorkbook(workbook);
console.log(JSON.stringify({
  sheets: workbook.SheetNames,
  bestSheetIndex: analysis.bestSheetIndex,
  recognizedSheets: analysis.recognizedSheets.map((sheet) => ({
    name: sheet.name,
    headerRowIndex: sheet.headerRowIndex,
    headerSpan: sheet.headerSpan,
    recognitionScore: sheet.recognitionScore,
    suggestions: sheet.suggestions.map((item) => ({ column: item.colIndex, header: item.header, field: item.field, confidence: item.confidence })),
    rows: sheet.rows.slice(Math.max(0, sheet.headerRowIndex - 1), sheet.dataStartIndex + 5),
  })),
}, null, 2));

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const JSZip = require("jszip");
const Codepage = require("codepage");
const NodeXLSX = require("xlsx-js-style");

const root = path.resolve(__dirname, "..");

function browserRuntime() {
  const context = { console, TextDecoder, TextEncoder, Uint8Array, ArrayBuffer, Promise, setTimeout, clearTimeout };
  context.self = context;
  context.window = context;
  vm.createContext(context);
  for (const relative of ["vendor/cpexcel.full.js", "vendor/xlsx.bundle.js", "vendor/jszip.min.js", "bom-core.js", "excel-reader.js"]) {
    vm.runInContext(fs.readFileSync(path.join(root, relative), "utf8"), context, { filename: relative });
  }
  return context;
}

async function readRows(runtime, bytes, filename) {
  const result = await runtime.ExcelReader.readWorkbook(bytes, { filename, encoding: "auto" });
  const sheet = result.workbook.Sheets[result.workbook.SheetNames[0]];
  return { result, rows: runtime.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false }) };
}

async function run() {
  const runtime = browserRuntime();
  assert.ok(runtime.cptable && runtime.cptable[936], "桌面/浏览器运行环境必须包含 CP936 中文代码页");

  const csvText = "\u7269\u6599\u7f16\u7801,\u7269\u6599\u540d\u79f0,\u6570\u91cf\r\nR-001,\u7535\u963b,2";
  const gbkCsv = Codepage.utils.encode(936, csvText);
  const csv = await readRows(runtime, gbkCsv, "\u771f\u5b9eGBK\u6599\u5355.csv");
  assert.equal(csv.result.encoding, "gb18030");
  assert.equal(csv.rows[0][0], "\u7269\u6599\u7f16\u7801");
  assert.equal(csv.rows[1][1], "\u7535\u963b");

  global.cptable = Codepage;
  const workbook = NodeXLSX.utils.book_new();
  NodeXLSX.utils.book_append_sheet(workbook, NodeXLSX.utils.aoa_to_sheet([
    ["\u7269\u6599\u7f16\u7801", "\u7269\u6599\u540d\u79f0", "\u6570\u91cf"],
    ["R-001", "\u7535\u963b", 2]
  ]), "BOM");

  const xlsBytes = NodeXLSX.write(workbook, { type: "buffer", bookType: "xls" });
  const xls = await readRows(runtime, xlsBytes, "\u771f\u5b9e\u4e2d\u6587.xls");
  assert.equal(xls.rows[0][0], "\u7269\u6599\u7f16\u7801");
  assert.equal(xls.rows[1][1], "\u7535\u963b");

  const standardXlsx = NodeXLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  const malformedZip = await JSZip.loadAsync(standardXlsx);
  const xmlNames = Object.keys(malformedZip.files).filter((name) => /^(?:xl\/.*|docProps\/.*)\.xml$/i.test(name));
  for (const name of xmlNames) {
    const xml = await malformedZip.file(name).async("string");
    malformedZip.file(name, Codepage.utils.encode(936, xml));
  }
  const malformedXlsx = await malformedZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const repaired = await readRows(runtime, malformedXlsx, "\u975e\u6807\u51c6GBK.xlsx");
  assert.equal(repaired.result.encoding, "gb18030", "内部为 GBK XML 的 XLSX 应自动选择中文修复结果");
  assert.equal(repaired.rows[0][0], "\u7269\u6599\u7f16\u7801");
  assert.equal(repaired.rows[1][1], "\u7535\u963b");

  console.log("browser runtime reads real GBK CSV, XLS and malformed XLSX without mojibake");
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

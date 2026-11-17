const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const XLSX = require("xlsx-js-style");
const Core = require("../bom-core");
const Reader = require("../excel-reader");
const Exporter = require("../template-exporter");
const Defaults = require("../default-templates");

const headers = ["No.", "Quantity", "Comment", "Designator", "Footprint", "Value",
  "Manufacturer Part", "Manufacturer", "Supplier Part", "Supplier"];
const data = [
  [1, 2, "12pF", "C1,C2", "C0402", "12pF", "MPN-12P", "Maker", "C1234", "LCSC"],
  [2, 1, "12pF", "C3", "C0402", "12pF", "MPN-12P", "Maker", "C1234", "LCSC"],
  [4, 1, "12pF", "C4", "C0402", "12pF", "MPN-12P", "Maker", "C9876", "LCSC"],
  [5, 1, "1.0K", "R1", "R0402", "1.0K", "", "", "", ""],
  [6, 1, "10K", "R2", "R0402", "10K", "", "", "", ""],
  [7, 1, "", "C5", "C0603", "100nF", "", "", "", ""],
  [8, 1, "12pF", "C6", "C0603", "12pF", "MPN-12P", "Maker", "C1234", "LCSC"],
  ["", "", "", "", "", "", "", "", "", ""]
];

function makeFile(rows, name) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "BOM");
  const bytes = XLSX.write(book, { type: "array", bookType: "xlsx" });
  return { name, size: bytes.byteLength, arrayBuffer: async () => bytes };
}

async function run() {
  const source = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8");
  const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
  assert.match(html, /data-mode="jlcpcb"/);
  assert.match(html, /data-mode="bom"/);
  const context = { console, BOMCore: Core, ExcelReader: Reader, XLSX, window: {},
    document: { querySelector: (key) => ({ value: key === "#multiplier" ? "3" : "" }) } };
  vm.runInNewContext(source.slice(0, source.lastIndexOf("\n  bindEvents();"))
    + "\n globalThis.ui = {state, parseDocument, availableOutputKeys};\n})();", context);
  const { state, parseDocument, availableOutputKeys } = context.ui;
  state.mode = "jlcpcb";
  assert.deepEqual([...availableOutputKeys()], Object.keys(Core.OUTPUT_DEFS));
  const doc = await parseDocument(makeFile([["导出信息"], headers, ...data], "嘉立创样例.xlsx"));
  assert.equal(doc.sheet.bomFormat, "jlcpcb");
  assert.equal(doc.sheet.headerRowIndex, 1);
  assert.deepEqual([...doc.mapping], ["", "quantity", "description", "designator", "footprint",
    "componentValue", "manufacturerPartNumber", "vendor", "supplierPartNumber", "supplierName"]);
  assert.deepEqual(Core.reanalyzeWithHeader(doc.sheet, 1, 1).suggestions.map(m => m.field), doc.mapping);
  assert.equal(doc.rows.length, 7, "序号间断和尾部空行不得影响有效行");
  const [first, , , noMpn, , noComment] = doc.rows;
  assert.deepEqual([first.model, first.description, first.vendor, first.partNumber, first.quantity],
    ["MPN-12P", "12pF", "Maker", "", 6]);
  assert.match(first.remark, /参数：12pF.*LCSC C1234/);
  assert.deepEqual([noMpn.model, noMpn.vendor, noMpn.quality, noMpn.manufacturerPartNumber], ["1.0K", "", "", ""]);
  assert.match(noMpn.remark, /封装：R0402.*厂家料号未填写/);
  assert.match(Core.revalidateRow(noMpn, 4).warnings.find(w => w.field === "model").message, /未填写厂家料号/);
  assert.equal(noComment.description, "100nF");
  const fullHeaders = headers.map(h => h === "Manufacturer Part" ? "Manufacturer Part Number" : h);
  const longHeaderDoc = await parseDocument(makeFile([fullHeaders, ...data], "长表头.xlsx"));
  assert.equal(longHeaderDoc.rows[0].model, "MPN-12P");
  const chinese = await parseDocument(makeFile([["序号", "数量", "名称", "位号", "封装", "数值", "厂家料号", "制造商", "商城编号", "供应商"], ...data], "中文表头.xlsx"));
  assert.equal(chinese.rows[0].model, "MPN-12P");
  assert.equal(chinese.rows[0].supplierPartNumber, "C1234");

  const result = await Exporter.buildFiles(doc.rows, { mode: "jlcpcb", multiplier: 3 }, {
    selectedOutputs: ["assembly", "procurement", "picking", "subcontract"],
    templates: Object.fromEntries(Object.entries(Defaults).map(([k, v]) => [k, { exactDefinition: v }]))
  }, Core, XLSX);
  assert.equal(result.files.length, 4);
  for (const key of ["assembly", "procurement", "picking", "subcontract"]) {
    assert.equal(result.outputs[key].length, 6, "仅合并同料号、同商城编号、同封装的行");
    assert.equal(result.outputs[key].reduce((s, row) => s + row.baseQuantity, 0), 8);
    assert.equal(result.outputs[key].reduce((s, row) => s + row.quantity, 0), 24);
  }
  const sheet = (key) => {
    const wb = XLSX.read(result.files.find(f => f.key === key).bytes, { type: "array", cellFormula: true });
    return wb.Sheets[wb.SheetNames[0]];
  };
  const a = sheet("assembly");
  assert.equal(a.C6.v, "MPN-12P");
  assert.equal(a.E6.v, 3);
  assert.equal(a.F6.f, "E6*$G$3");
  assert.equal(a.H6.v, "C0402");
  const p = sheet("procurement");
  assert.equal(p.C5.v, "MPN-12P");
  assert.equal(p.G5.v, "Maker");
  assert.match(p.M5.v.replace(/\s/g, ""), /C1234/);
  assert.equal(p.J5.f, "F5*$J$4");
  assert.equal(p.J5.v, 9);
  assert.equal(p.K5.f, "J5+I5");

  state.mode = "bom";
  const ad = await parseDocument(makeFile([["Comment", "Value", "Designator", "Footprint", "Quantity", "Vendor"],
    ["电阻", "AD-MODEL", "R1", "R0402", 2, "Maker"]], "AD.xlsx"));
  assert.notEqual(ad.sheet.bomFormat, "jlcpcb");
  assert.equal(ad.rows[0].model, "AD-MODEL");
  assert.equal(ad.rows[0].vendor, "Maker");
  assert.equal(ad.rows[0].quantity, 6);
  console.log("嘉立创 BOM 导入、公司模板导出、数量公式和 AD 回归通过");
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

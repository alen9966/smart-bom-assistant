const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const XLSX = require("xlsx-js-style");
const BOMCore = require("../bom-core.js");
const ExcelReader = require("../excel-reader.js");
const Exporter = require("../template-exporter.js");
const defaults = require("../default-templates.js");

function element() {
  return {
    value: "", innerHTML: "", dataset: {}, style: {},
    events: {}, classList: { add() {}, remove() {}, toggle() {} },
    setAttribute() {}, querySelectorAll() { return []; },
    addEventListener(type, handler) { this.events[type] = handler; }
  };
}

async function run() {
  const elements = new Map();
  const get = (id) => {
    if (!elements.has(id)) elements.set(id, element());
    return elements.get(id);
  };
  get("#multiplier").value = "3";
  const app = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
  const init = app.lastIndexOf("\n  bindEvents();");
  assert.ok(init > 0);
  const context = { document: { querySelector: get, querySelectorAll: () => [] }, window: {},
    console, BOMCore, ExcelReader, XLSX, setTimeout: () => 0, clearTimeout() {} };
  vm.runInNewContext(app.slice(0, init) + `
    globalThis.appTest = {state, parseDocument, renderBomFileList, activateDocument,
      reloadActiveDocumentWithEncoding, documentMetadata, combinedDocumentMetadata};
  })();`, context);
  const ui = context.appTest;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["名称", "型号", "位号", "数量"], ["电容", "0603X106M250NT", "C1", 2]
  ]), "BOM");
  const bytes = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const first = await ui.parseDocument({name: "ZC7.820.0100-V1.xlsx", size: bytes.byteLength, arrayBuffer: async () => bytes});
  const second = await ui.parseDocument({name: "子板.xlsx", size: bytes.byteLength, arrayBuffer: async () => bytes});
  assert.equal(first.pcbRevision, "ZC7.820.0100-V1");
  assert.equal(second.pcbRevision, "");
  ui.state.documents.push(first, second);
  const fields = [first, second].map(doc => Object.assign(element(), { dataset: { pcbDocumentId: doc.id } }));
  get("#bomFileList").querySelectorAll = selector => selector === "[data-pcb-document-id]" ? fields : [];
  ui.activateDocument(first.id);
  assert.equal([...get("#bomFileList").innerHTML.matchAll(/data-pcb-document-id=/g)].length, 2);
  fields[0].value = " ZC7.820.0100-V2 ";
  fields[0].events.input();
  fields[1].value = "ZC7.820.0200-V3";
  fields[1].events.input();
  ui.activateDocument(second.id);
  ui.activateDocument(first.id);
  get("#encodingSelect").value = "utf8";
  await ui.reloadActiveDocumentWithEncoding();
  const [a, b] = ui.state.documents;
  assert.equal(a.pcbRevision.trim(), "ZC7.820.0100-V2", "重读后应保留该 BOM 的手动版号");
  assert.equal(a.encodingPreference, "utf8", "重读结果不能被旧状态覆盖");
  assert.equal(b.pcbRevision, "ZC7.820.0200-V3");

  const metadata = { projectCode: "TEST", mode: "bom", multiplier: 3, pcbVendor: "测试板厂" };
  const options = { selectedOutputs: ["assembly", "procurement"], skipInvalid: true, purchaseMode: "auto",
    templates: Object.fromEntries(Object.entries(defaults).map(([key, definition]) => [key, { exactDefinition: definition }])) };
  const revisionA = "ZC7.820.0100-V2";
  const revisionB = "ZC7.820.0200-V3";
  for (const [doc, revision] of [[a, revisionA], [b, revisionB]]) {
    const meta = ui.documentMetadata(doc, metadata);
    const result = await Exporter.buildFiles(doc.rows, meta, options, BOMCore, XLSX);
    const assembly = XLSX.read(result.files.find(file => file.key === "assembly").bytes, { type: "array" });
    assert.equal(assembly.Sheets[assembly.SheetNames[0]].C1.v, revision);
    assert.deepEqual(result.outputs.procurement.filter(BOMCore.isPrintedBoard).map(row => row.model), [revision]);
  }
  const combinedMeta = ui.combinedDocumentMetadata([a, b, { ...a }], metadata);
  assert.deepEqual([...combinedMeta.pcbRevisions], [revisionA, revisionB]);
  const combined = await Exporter.buildFiles([...a.rows, ...b.rows], combinedMeta, options, BOMCore, XLSX);
  assert.deepEqual(combined.outputs.procurement.filter(BOMCore.isPrintedBoard).map(row => [row.model, row.quantity]),
    [[revisionA, 3], [revisionB, 3]]);
  const assembly = XLSX.read(combined.files.find(file => file.key === "assembly").bytes, { type: "array", cellStyles: true });
  const sheet = assembly.Sheets[assembly.SheetNames[0]];
  assert.equal(sheet.C1.v, `${revisionA}\n${revisionB}`);
  assert.ok(sheet["!rows"][0].hpt > 20, "合并清单表头应容纳多个版号");

  fields[0].value = "";
  fields[0].events.input();
  const cleared = BOMCore.buildOutputs(a.rows, ui.documentMetadata(a, metadata));
  assert.equal(cleared.procurement.filter(BOMCore.isPrintedBoard).length, 0, "清空后不得从文件名再次补入旧版号");
  console.log("per-BOM revision editing, encoding reload and separate/combined exports passed");
}

run().catch(error => { console.error(error); process.exitCode = 1; });

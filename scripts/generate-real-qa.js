const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx-js-style");
const Core = require("../bom-core.js");
const Exporter = require("../template-exporter.js");
const Definitions = require("../default-templates.js");

async function run() {
  const source = process.argv[2];
  const outputDir = process.argv[3] || path.join(__dirname, "..", ".qa-default-templates", "real-v17");
  const workbook = XLSX.readFile(source);
  const analysis = Core.analyzeWorkbook(workbook);
  const sheet = analysis.recognizedSheets[0];
  const mapping = sheet.suggestions.map((item) => item.field || "");
  const extracted = Core.extractRows(sheet, mapping, 4);
  extracted.rows.forEach((row) => { row.sourceFile = path.basename(source); });
  const templates = Object.fromEntries(Object.entries(Definitions).map(([key, exactDefinition]) => [key, { exactDefinition }]));
  const metadata = { projectCode: "ZCJY260704", productModel: "共视接收机", projectName: "整机采购", department: "设备与系统", batch: "2608", multiplier: 4, sourceFile: path.basename(source), mode: "bom" };
  const result = await Exporter.buildFiles(extracted.rows, metadata, { selectedOutputs: ["assembly", "picking", "procurement", "subcontract"], skipInvalid: true, purchaseMode: "auto", templates }, Core, XLSX);
  fs.mkdirSync(outputDir, { recursive: true });
  for (const file of result.files) fs.writeFileSync(path.join(outputDir, file.filename), file.bytes);
  console.log(JSON.stringify({ outputDir, rows: extracted.rows.length, firstRows: extracted.rows.slice(0, 5).map((row) => ({ category: row.category, description: row.description, model: row.model, footprint: row.footprint })), files: result.files.map((file) => file.filename) }, null, 2));
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

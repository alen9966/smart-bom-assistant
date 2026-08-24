const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx-js-style");
const Core = require("../bom-core.js");
const defaults = require("../default-templates.js");

const rows = [
  { partNumber: "R-001", description: "电阻", model: "10kΩ 1%", quality: "I", designator: "R1, R2", baseQuantity: 2, quantityRaw: "2", quantity: 8, unit: "只", vendor: "风华", source: "外购", drawingNo: "", assemblyName: "主板", remark: "电装", sourceFile: "验证BOM.xlsx", sourceSheet: "BOM", sourceRow: 7, errors: [], warnings: [] },
  { partNumber: "C-001", description: "电容", model: "100nF/50V", quality: "I", designator: "C1", baseQuantity: 4, quantityRaw: "4", quantity: 16, unit: "只", vendor: "国巨", source: "外购", drawingNo: "", assemblyName: "主板", remark: "电装", sourceFile: "验证BOM.xlsx", sourceSheet: "BOM", sourceRow: 8, errors: [], warnings: [] },
];
const templates = Object.fromEntries(Object.entries(defaults).map(([key, definition]) => [key, {
  workbook: XLSX.read(Buffer.from(definition.base64, "base64"), { type: "buffer", cellStyles: true, cellNF: true }),
  inspection: definition.inspection,
}]));
const result = Core.buildWorkbook(rows, {
  projectCode: "ZCJY260704", productModel: "共视接收机", projectName: "整机采购", department: "设备与系统",
  batch: "2608", multiplier: 4, sourceFile: "验证BOM.xlsx", mode: "bom",
}, { selectedOutputs: ["assembly", "picking", "procurement"], skipInvalid: true, purchaseMode: "auto", templates });
const outputDir = path.resolve(__dirname, "..", ".qa-default-templates");
fs.mkdirSync(outputDir, { recursive: true });
XLSX.writeFile(result.workbook, path.join(outputDir, "generated-defaults.xlsx"), { compression: true });

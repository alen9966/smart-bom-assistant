const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx-js-style");
const Core = require("../bom-core.js");
const Codepage = require("codepage");
const ExcelReader = require("../excel-reader.js");
const DefaultTemplates = require("../default-templates.js");

const workbook = XLSX.utils.book_new();
const source = [
  ["项目：智能识别测试"],
  ["制表日期", "2026-08-10"],
  ["序号", "物料编号", "参考位号", "元件名称", "规格型号", "每套用量", "计量单位", "供应厂家", "采购属性", "Variant"],
  [1, "R-001", "R1, R2", "电阻", "10kΩ 1%", 2, "只", "厚声", "外购", "整机采购"],
  [2, "C-001", "C1", "电容", "100nF/50V", "4 pcs", "只", "国巨", "标准件", "整机采购"],
  [3, "U-001", "U1", "控制芯片", "MCU-A", "数量待定", "只", "示例厂家", "外购", "底板采购"]
];
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(source), "主BOM");

const analysis = Core.analyzeWorkbook(workbook);
assert.equal(analysis.hasRecognizedSheet, true, "应识别 BOM 工作表");
const sheet = analysis.sheets[0];
assert.equal(sheet.headerRowIndex, 2, "应定位到第 3 行表头");
assert.ok(sheet.matchedCount >= 6, "应匹配多数关键字段");

const mapping = sheet.suggestions.map((item) => item.field || "");
const mappingCheck = Core.validateMapping(mapping);
assert.equal(mappingCheck.valid, true, "自动映射应包含数量与物料标识");

const extracted = Core.extractRows(sheet, mapping, 4);
assert.equal(extracted.rows.length, 3, "应读取 3 行物料");
assert.equal(extracted.rows[0].quantity, 8, "数量应乘以生产数量");
assert.equal(extracted.rows[0].baseQuantity, 2, "应保留未放大的原始数量");
assert.equal(extracted.rows[2].status, "error", "模糊数量应明确报错");
assert.match(extracted.rows[2].errors[0].message, /数量待定/, "报错应包含原始值");

const outputs = Core.buildOutputs(extracted.rows, { skipInvalid: true, purchaseMode: "auto" });
assert.equal(outputs.picking.length, 2, "领料汇总应跳过错误行");
assert.equal(outputs.procurement.length, 2, "外购物料应进入采购清单");
assert.ok(outputs.issues.length >= 1, "问题应进入问题清单");

const pcbOutputs = Core.buildOutputs(extracted.rows, {
  skipInvalid: true,
  purchaseMode: "auto",
  multiplier: 4,
  pcbRevision: "ZC7.820.0155-V3",
  pcbVendor: "嘉立创"
});
const pcbRow = pcbOutputs.procurement.find((row) => row.category === "印制板" || row.model === "ZC7.820.0155-V3");
assert.ok(pcbRow, "填写 PCB 版号后采购清单应追加印制板");
assert.equal(pcbRow.model, "ZC7.820.0155-V3", "印制板名称应为版号");
assert.equal(pcbRow.description, "ZC7.820.0155-V3", "印制板名称应为版号");
assert.equal(pcbRow.vendor, "嘉立创", "印制板生产厂家应使用页面 PCB 板厂");
assert.equal(pcbRow.baseQuantity, 1, "印制板单机数量默认为 1");
assert.equal(pcbRow.quantity, 4, "印制板采购数量应等于生产数量");
assert.equal(pcbOutputs.picking.some((row) => row.category === "印制板"), false, "印制板只进入采购清单，不进入领料单");
assert.equal(pcbOutputs.procurement.filter((row) => row.model === "ZC7.820.0155-V3").length, 1, "同一版号不得重复追加印制板");

const assemblyExtracted = Core.extractRows(sheet, mapping, 1);
const mechanicalRow = { ...assemblyExtracted.rows[0], partNumber: "M-001", designator: "M1", description: "安装螺钉", model: "M3", quantityRaw: "3", baseQuantity: 3, quantity: 3, errors: [], warnings: [], status: "ok" };
const assemblyOutputs = Core.buildOutputs([...assemblyExtracted.rows, mechanicalRow], { skipInvalid: true, purchaseMode: "auto", multiplier: 4 });
assert.equal(assemblyOutputs.picking[0].quantity, 2, "装配文件领料单应保持原数量，不乘生产数量");
assert.equal(assemblyOutputs.subcontract.length, 2, "外协备料清单应只保留电阻和电容");
assert.equal(assemblyOutputs.subcontract.find((row) => row.partNumber === "R-001").quantity, 8, "外协电阻应乘生产数量");
assert.equal(assemblyOutputs.subcontract.find((row) => row.partNumber === "C-001").quantity, 16, "外协电容应乘生产数量");
assert.equal(assemblyOutputs.subcontract.find((row) => row.partNumber === "R-001").baseQuantity, 2, "外协阻容单机数量应保留汇总后的原数量");

const bomOutputsWithSubcontract = Core.buildOutputs(extracted.rows, { skipInvalid: true, purchaseMode: "auto", multiplier: 4 });
assert.equal(bomOutputsWithSubcontract.subcontract.find((row) => row.partNumber === "R-001").quantity, 8, "BOM 已放大的数量在外协清单中不得重复相乘");

const subcontractOnly = Core.buildWorkbook(assemblyExtracted.rows, {
  projectCode: "TEST-OUTSOURCE",
  projectName: "外协备料测试",
  multiplier: 4,
  sourceFile: "装配清单.xlsx",
  mode: "assembly"
}, {
  selectedOutputs: ["subcontract"],
  skipInvalid: true,
  purchaseMode: "auto"
});
assert.deepEqual(subcontractOnly.workbook.SheetNames, ["导出说明", "外协阻容备料清单", "数据问题清单"]);
assert.equal(subcontractOnly.workbook.Sheets["外协阻容备料清单"].H7.v, 8, "外协备料导出数量应为原数量乘生产数量");
const subcontractPath = path.join(__dirname, "tmp", "subcontract-output.xlsx");
fs.mkdirSync(path.dirname(subcontractPath), { recursive: true });
XLSX.writeFile(subcontractOnly.workbook, subcontractPath, { compression: true });

const built = Core.buildWorkbook(extracted.rows, {
  projectCode: "TEST-001",
  projectName: "智能识别测试",
  multiplier: 4,
  sourceFile: "测试BOM.xlsx",
  mode: "bom"
}, {
  selectedOutputs: ["assembly", "picking", "procurement", "detail", "purchased"],
  skipInvalid: true,
  purchaseMode: "auto"
});
assert.deepEqual(built.workbook.SheetNames, ["导出说明", "装配清单", "领料单", "采购清单", "明细表", "外购件汇总表", "数据问题清单"]);

const outputDir = path.join(__dirname, "tmp");
fs.mkdirSync(outputDir, { recursive: true });
const outputPath = path.join(outputDir, "test-output.xlsx");
XLSX.writeFile(built.workbook, outputPath, { compression: true });
assert.ok(fs.statSync(outputPath).size > 5000, "导出 Excel 应包含有效内容");

const reread = XLSX.readFile(outputPath);
assert.equal(reread.SheetNames.includes("采购清单"), true, "导出文件应可再次读取");

const templateWorkbook = XLSX.utils.book_new();
const templateSheet = XLSX.utils.aoa_to_sheet([
  ["采购申请单 {{项目工号}}"],
  ["项目名称", ""],
  [],
  ["项次", "物料编码", "物料名称", "采购数量", "单位", "供应商"],
  ["", "", "", "", "", ""]
]);
templateSheet.A5.s = { fill: { fgColor: { rgb: "FFF2CC" } }, font: { name: "Microsoft YaHei" } };
templateSheet.B5.s = { fill: { fgColor: { rgb: "FFF2CC" } } };
templateSheet.C5.s = { fill: { fgColor: { rgb: "FFF2CC" } } };
templateSheet.D5.s = { fill: { fgColor: { rgb: "FFF2CC" } }, numFmt: "0.00" };
templateSheet.E5.s = { fill: { fgColor: { rgb: "FFF2CC" } } };
templateSheet.F5.s = { fill: { fgColor: { rgb: "FFF2CC" } } };
XLSX.utils.book_append_sheet(templateWorkbook, templateSheet, "公司采购模板");
const inspection = Core.inspectTemplate(templateWorkbook, "procurement");
assert.equal(inspection.valid, true, "应自动找到采购模板表头");
assert.equal(inspection.headerRowIndex, 3, "应定位模板第 4 行表头");
assert.equal(Core.inspectTemplate(templateWorkbook, "subcontract").valid, true, "外协备料清单也应支持独立模板");
assert.equal(Core.OUTPUT_DEFS.assembly.primary, true, "装配清单应标记为常用输出");
assert.equal(Core.OUTPUT_DEFS.picking.primary, undefined, "领料单不应标记为常用输出");
assert.equal(Core.OUTPUT_DEFS.procurement.primary, true, "采购清单应标记为常用输出");
assert.equal(Core.OUTPUT_DEFS.subcontract.primary, true, "外协阻容备料清单应标记为常用输出");

function simpleTemplate(sheetName, quantityHeader) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    [`${sheetName} {{项目工号}}`],
    [],
    ["序号", "物料编码", "物料名称", quantityHeader, "单位"],
    ["", "", "", "", ""]
  ]), sheetName);
  return book;
}

const commonTemplates = {
  assembly: simpleTemplate("公司装配模板", "数量"),
  picking: simpleTemplate("公司领料模板", "领料数量"),
  procurement: templateWorkbook
};
const commonInspections = Object.fromEntries(Object.entries(commonTemplates).map(([key, book]) => [key, Core.inspectTemplate(book, key)]));
assert.equal(commonInspections.assembly.valid, true, "装配清单应支持独立模板");
assert.equal(commonInspections.picking.valid, true, "领料单应支持独立模板");
assert.equal(commonInspections.procurement.valid, true, "采购清单应支持独立模板");

const templated = Core.buildWorkbook(extracted.rows, {
  projectCode: "TEST-001",
  projectName: "模板测试项目",
  multiplier: 4,
  sourceFile: "测试BOM.xlsx",
  mode: "bom"
}, {
  selectedOutputs: ["procurement"],
  skipInvalid: true,
  purchaseMode: "auto",
  templates: { procurement: { workbook: templateWorkbook, inspection } }
});
const templatedSheet = templated.workbook.Sheets["采购清单"];
assert.equal(templatedSheet.A1.v, "采购申请单 TEST-001", "应替换模板占位符");
assert.equal(templatedSheet.B2.v, "模板测试项目", "应填写标签右侧的项目元数据");
assert.equal(templatedSheet.A5.v, 1, "应从模板表头下一行开始写序号");
assert.equal(templatedSheet.B5.v, "R-001", "应按模板列名写入物料编码");
assert.equal(templatedSheet.D5.v, 8, "应按模板列名写入采购数量");
assert.equal(templatedSheet.A5.s.fill.fgColor.rgb, "FFF2CC", "应保留模板数据行样式");
const templatedPath = path.join(outputDir, "templated-output.xlsx");
XLSX.writeFile(templated.workbook, templatedPath, { compression: true });
const templatedReread = XLSX.readFile(templatedPath, { cellStyles: true });
assert.equal(templatedReread.Sheets["采购清单"].B5.v, "R-001", "模板输出文件应可再次读取");

const commonTemplated = Core.buildWorkbook(extracted.rows, {
  projectCode: "COMMON-001",
  multiplier: 4,
  sourceFile: "测试BOM.xlsx",
  mode: "bom"
}, {
  selectedOutputs: ["assembly", "picking", "procurement"],
  skipInvalid: true,
  purchaseMode: "auto",
  templates: Object.fromEntries(Object.entries(commonTemplates).map(([key, book]) => [key, { workbook: book, inspection: commonInspections[key] }]))
});
assert.deepEqual(commonTemplated.workbook.SheetNames, ["导出说明", "装配清单", "领料单", "采购清单", "数据问题清单"], "常用三类清单应能同时使用各自模板生成");
assert.equal(commonTemplated.workbook.Sheets["装配清单"].B4.v, "R-001", "装配模板应填入 BOM 数据");
assert.equal(commonTemplated.workbook.Sheets["领料单"].D4.v, 8, "领料模板应填入汇总数量");
assert.equal(commonTemplated.workbook.Sheets["采购清单"].D5.v, 8, "采购模板应填入采购数量");

const builtInTemplates = Object.fromEntries(Object.entries(DefaultTemplates).map(([key, definition]) => [key, {
  workbook: XLSX.read(Buffer.from(definition.base64, "base64"), { type: "buffer", cellStyles: true, cellFormula: true, cellNF: true }),
  inspection: definition.inspection
}]));
const builtInResult = Core.buildWorkbook(extracted.rows, {
  projectCode: "ZCJY-DEFAULT",
  productModel: "共视接收机",
  projectName: "默认模板验证",
  department: "设备与系统",
  batch: "2608",
  multiplier: 4,
  sourceFile: "源BOM.xlsx",
  mode: "bom"
}, {
  selectedOutputs: ["assembly", "picking", "procurement"],
  skipInvalid: true,
  purchaseMode: "auto",
  templates: builtInTemplates
});
const defaultAssembly = builtInResult.workbook.Sheets["装配清单"];
const defaultPicking = builtInResult.workbook.Sheets["领料单"];
const defaultProcurement = builtInResult.workbook.Sheets["采购清单"];
assert.equal(defaultAssembly.B6.v, "电阻", "公司装配默认模板应从第 6 行填入数据");
assert.equal(defaultAssembly.E6.v, 2, "公司装配默认模板应保持单套数量");
assert.equal(defaultPicking.C4.v, "10kΩ 1%", "公司领料默认模板应从第 4 行填入名称型号");
assert.equal(defaultPicking.E4.v, 2, "公司领料默认模板应填入单机数量");
assert.equal(defaultPicking.G4.v, 8, "公司领料默认模板应填入领料总数量");
assert.equal(defaultPicking.H4.v, "2608", "公司领料默认模板应填入批次");
assert.equal(defaultProcurement.C4.v, "10kΩ 1%", "公司采购默认模板应从第 4 行填入名称型号");
assert.equal(defaultProcurement.F4.v, 2, "公司采购默认模板应填入单机数量");
assert.equal(defaultProcurement.J4.v, 8, "公司采购默认模板应填入装机总数");
assert.equal(defaultProcurement.K4.v, 8, "公司采购默认模板应填入采购总数");
const builtInPath = path.join(outputDir, "built-in-default-output.xlsx");
XLSX.writeFile(builtInResult.workbook, builtInPath, { compression: true });
assert.equal(XLSX.readFile(builtInPath).SheetNames.includes("装配清单"), true, "公司默认模板输出应可再次读取");

const firstBomRow = { ...extracted.rows[0], sourceFile: "BOM-A.xlsx" };
const secondBomRow = { ...extracted.rows[0], quantity: 12, sourceFile: "BOM-B.xlsx", sourceRow: 9 };
const combined = Core.buildOutputs([firstBomRow, secondBomRow], { skipInvalid: true, purchaseMode: "auto" });
assert.equal(combined.picking.length, 1, "相同物料跨 BOM 合并后应为一行");
assert.equal(combined.picking[0].quantity, 20, "跨 BOM 数量应正确相加");
assert.match(combined.picking[0].traceSummary, /BOM-A\.xlsx/, "合并结果应保留第一个 BOM 来源");
assert.match(combined.picking[0].traceSummary, /BOM-B\.xlsx/, "合并结果应保留第二个 BOM 来源");

const passiveSheet = XLSX.utils.aoa_to_sheet([
  ["Class", "Designator", "Value", "Quality", "Quantity", "Vendor", "Variant"],
  ["电容", "C1", "1206B102K202NT", "I", 1, "风华", "整机采购"],
  ["电阻", "R1", "RC-02W22R0FT", "I", 2, "风华", "整机采购"],
]);
const passiveAnalysis = Core.analyzeSheet(passiveSheet, "AD BOM", 0);
const passiveMapping = passiveAnalysis.suggestions.map((item) => item.field || "");
const passiveRows = Core.extractRows(passiveAnalysis, passiveMapping, 4).rows;
assert.equal(passiveRows[0].description, "1nF/2kV", "缺少 Comment 时应从电容型号推导名称");
assert.equal(passiveRows[0].footprint, "1206C", "缺少 Footprint 时应从电容型号推导封装");
assert.equal(passiveRows[1].description, "22Ω", "缺少 Comment 时应从电阻型号推导名称");
assert.equal(passiveRows[1].footprint, "0402R", "缺少 Footprint 时应从电阻型号推导封装");

const headerChecks = Core.headerRequirementStatus(passiveMapping, "bom", ["assembly", "picking", "procurement"]);
assert.equal(headerChecks.find((item) => item.key === "quantity").present, true, "已映射数量时表头检查应为绿色");
assert.equal(headerChecks.find((item) => item.key === "footprint").present, false, "缺少封装表头时检查应为红色");
assert.match(headerChecks.find((item) => item.key === "footprint").impact, /装配清单.*封装/, "缺失表头应说明具体影响");
assert.deepEqual(Core.headerRequirementStatus(passiveMapping, "bom", ["procurement"]).some((item) => item.key === "footprint"), false, "只选择采购清单时不应提示装配专用封装表头");
assert.deepEqual(Core.headerRequirementStatus(passiveMapping, "bom", []), [], "未勾选清单时不应提示任何表头要求");

const subcontractTemplateMatches = {
  subcontract: DefaultTemplates.subcontract.inspection.matches.map((match) => ({ key: match.key }))
};
const subcontractColumns = Core.templateColumnStatus(passiveMapping, passiveRows, ["subcontract"], subcontractTemplateMatches);
assert.deepEqual(subcontractColumns.some((item) => item.key === "partNumber"), false, "外协模板不含物料编码时不得要求物料编码");
assert.deepEqual(subcontractColumns.some((item) => item.key === "unit"), false, "外协模板不含单位时不得要求单位");
assert.deepEqual(subcontractColumns.some((item) => item.key === "manufacturerPartNumber"), false, "外协模板不含厂家料号时不得要求厂家料号");
assert.equal(subcontractColumns.find((item) => item.key === "modelOrDescription").status, "derived", "型号/名称可由 Value 等现有字段生成");
assert.equal(subcontractColumns.find((item) => item.key === "baseQuantity").status, "derived", "单机数量应由 Quantity 解析生成");
assert.equal(subcontractColumns.find((item) => item.key === "quantity").status, "present", "Quantity 已识别时总数量来源应显示已识别");
assert.equal(subcontractColumns.filter((item) => item.key === "quantity").length, 2, "模板中两个数量列应逐列显示，不得合并");
assert.equal(subcontractColumns.filter((item) => item.status === "missing").length, 0, "截图对应的外协模板字段不应误报缺失");

const missingQuantityColumns = Core.templateColumnStatus(passiveMapping.filter((key) => key !== "quantity"), [], ["subcontract"], subcontractTemplateMatches);
assert.equal(missingQuantityColumns.find((item) => item.key === "baseQuantity").status, "missing", "确实没有数量时应提示缺少");
if (!process.env.KEEP_QA) {
  fs.rmSync(outputPath, { force: true });
  fs.rmSync(templatedPath, { force: true });
  fs.rmSync(subcontractPath, { force: true });
  fs.rmSync(builtInPath, { force: true });
}
function exactArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function runEncodingTests() {
  const csvText = "物料编码,物料名称,数量\r\nR-001,电阻,2\r\nC-001,电容,4";
  const gbkBuffer = Codepage.utils.encode(936, csvText);
  const gbkResult = await ExcelReader.readWorkbook(exactArrayBuffer(gbkBuffer), { filename: "伪装成Excel.xlsx", encoding: "auto" });
  const gbkRows = XLSX.utils.sheet_to_json(gbkResult.workbook.Sheets[gbkResult.workbook.SheetNames[0]], { header: 1, raw: false });
  assert.equal(gbkResult.kind, "text", "扩展名为 xlsx 时也应识别真实文本格式");
  assert.equal(gbkResult.encoding, "gb18030", "GBK 文件应自动选择简体中文编码");
  assert.equal(gbkRows[0][0], "物料编码", "GBK 表头不得乱码");
  assert.equal(gbkRows[1][1], "电阻", "GBK 数据不得乱码");

  const utf8Buffer = new TextEncoder().encode(csvText);
  const utf8Result = await ExcelReader.readWorkbook(utf8Buffer, { filename: "标准.csv", encoding: "auto" });
  assert.equal(utf8Result.encoding, "utf8", "UTF-8 CSV 应保持 UTF-8");
  assert.equal(XLSX.utils.sheet_to_json(utf8Result.workbook.Sheets[utf8Result.workbook.SheetNames[0]], { header: 1 })[0][1], "物料名称");

  const xmlText = `<?xml version="1.0" encoding="gb2312"?>
  <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
    <Worksheet ss:Name="Sheet1"><Table>
      <Row><Cell><Data ss:Type="String">物料编码</Data></Cell><Cell><Data ss:Type="String">物料名称</Data></Cell><Cell><Data ss:Type="String">数量</Data></Cell></Row>
      <Row><Cell><Data ss:Type="String">R-001</Data></Cell><Cell><Data ss:Type="String">电阻</Data></Cell><Cell><Data ss:Type="Number">2</Data></Cell></Row>
    </Table></Worksheet>
  </Workbook>`;
  const xmlGbk = Codepage.utils.encode(936, xmlText);
  const xmlResult = await ExcelReader.readWorkbook(exactArrayBuffer(xmlGbk), { filename: "系统导出.xlsx", encoding: "auto" });
  assert.equal(xmlResult.kind, "xml", "应识别伪装成 xlsx 的 XML 表格");
  assert.equal(xmlResult.encoding, "gb18030", "GBK XML 表格应自动纠正编码");
  assert.equal(XLSX.utils.sheet_to_json(xmlResult.workbook.Sheets[xmlResult.workbook.SheetNames[0]], { header: 1 })[1][1], "电阻");

  const normalWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(normalWorkbook, XLSX.utils.aoa_to_sheet([["物料编码", "物料名称", "数量"], ["R-001", "电阻", 2]]), "BOM");
  const normalBytes = XLSX.write(normalWorkbook, { type: "array", bookType: "xlsx" });
  const normalResult = await ExcelReader.readWorkbook(normalBytes, { filename: "标准.xlsx", encoding: "auto" });
  assert.equal(normalResult.kind, "xlsx", "标准 xlsx 应识别为 ZIP-XLSX");
  assert.equal(normalResult.encoding, "utf8", "标准 xlsx 不应被错误地改成 GBK");
  assert.equal(normalResult.suspicious, false, "正常中文不应被判为乱码");
}

runEncodingTests()
  .then(() => console.log("core and encoding tests passed"))
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });

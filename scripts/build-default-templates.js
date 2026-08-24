const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const configs = {
  assembly: {
    filename: "装配清单公司默认模板.xlsx",
    source: "default-assembly.xlsx",
    inspection: {
      valid: true,
      outputKey: "assembly",
      sheetName: "ZC7.820.0040-V3_装配清单",
      headerRowIndex: 4,
      dataStartIndex: 5,
      dataEndIndex: 122,
      dataMaxCol: 8,
      sheetPath: "xl/worksheets/sheet1.xml",
      openView: "normal",
      printAreaMaxCol: "I",
      fitToPageWidth: true,
      metadataCells: { C1: "{{产品型号}}", C2: "{{生成时间}}", G2: "{{装配变量}}", C3: "{{项目工号}} / {{项目名称}}", G3: "{{生产数量}}" },
      quantityFormula: { baseColIndex: 4, multiplierCell: "G3" },
      matches: [
        [0, "_index"], [1, "description"], [2, "model"], [3, "designator"],
        [4, "baseQuantity"], [5, "quantity"], [6, "vendor"], [7, "footprint"], [8, "remark"],
      ],
    },
  },
  picking: {
    filename: "领料单公司默认模板.xlsx",
    source: "default-picking.xlsx",
    inspection: {
      valid: true,
      outputKey: "picking",
      sheetName: "Sheet1",
      headerRowIndex: 1,
      dataStartIndex: 3,
      dataEndIndex: 31,
      footerStartIndex: 32,
      dataMaxCol: 10,
      categorize: true,
      rowStyleRows: { category: 6, item: 5 },
      styleOverrides: { category: { 2: "19" } },
      openView: "pageBreakPreview",
      printAreaMaxCol: "K",
      printTitleRows: "$1:$3",
      pageFooter: "&C第 &P 页，共 &N 页",
      sheetPath: "xl/worksheets/sheet1.xml",
      metadataCells: { A1: "领用部门：{{使用部门}} 工号：{{项目工号}} 产品型号名称：{{产品型号}} {{项目名称}} 投产数量：{{生产数量}} 筛选要求：无" },
      matches: [
        [0, "_index"], [1, "drawingNo"], [2, "modelOrDescription"], [3, "quality"],
        [4, "baseQuantity"], [5, "vendor"], [6, "quantity"], [7, "_batch"], [10, "_electricalAssembly"],
      ],
    },
  },
  procurement: {
    filename: "采购清单公司默认模板.xlsx",
    source: "default-procurement.xlsx",
    inspection: {
      valid: true,
      outputKey: "procurement",
      sheetName: "Sheet1",
      headerRowIndex: 2,
      dataStartIndex: 3,
      dataEndIndex: 90,
      footerStartIndex: 91,
      dataMaxCol: 15,
      categorize: true,
      rowStyleRows: { category: 4, item: 5 },
      openView: "pageBreakPreview",
      printAreaMaxCol: "P",
      printTitleRows: "$1:$3",
      pageFooter: "&C第 &P 页，共 &N 页",
      fitToPageWidth: true,
      sheetPath: "xl/worksheets/sheet1.xml",
      metadataCells: { A2: "申请部门：{{使用部门}} 工号：{{项目工号}} 产品型号名称：{{产品型号}} {{项目名称}} 投产数量：{{生产数量}} 筛选要求：无" },
      matches: [
        [0, "_index"], [1, "drawingNo"], [2, "modelOrDescription"], [3, "quality"],
        [5, "baseQuantity"], [6, "vendor"], [7, "_yes"], [8, "_zero"], [9, "quantity"],
        [10, "quantity"], [12, "remark"],
      ],
    },
  },
  subcontract: {
    filename: "外协阻容备料清单公司默认模板.xlsx",
    source: "default-subcontract.xlsx",
    inspection: {
      valid: true,
      outputKey: "subcontract",
      sheetName: "Sheet1",
      headerRowIndex: 2,
      dataStartIndex: 3,
      dataEndIndex: 90,
      footerStartIndex: 91,
      dataMaxCol: 15,
      categorize: true,
      categoryKey: "componentType",
      rowStyleRows: { category: 4, item: 5 },
      openView: "pageBreakPreview",
      printAreaMaxCol: "P",
      printTitleRows: "$1:$3",
      pageFooter: "&C第 &P 页，共 &N 页",
      fitToPageWidth: true,
      sheetPath: "xl/worksheets/sheet1.xml",
      metadataCells: { A2: "申请部门：{{使用部门}} 工号：{{项目工号}} 产品型号名称：{{产品型号}} {{项目名称}} 投产数量：{{生产数量}} 筛选要求：无" },
      matches: [
        [0, "_index"], [1, "drawingNo"], [2, "modelOrDescription"], [3, "quality"],
        [5, "baseQuantity"], [6, "vendor"], [7, "_yes"], [8, "_zero"], [9, "quantity"],
        [10, "quantity"], [12, "remark"],
      ],
    },
  },
};

const payload = {};
for (const [key, config] of Object.entries(configs)) {
  const inspection = { ...config.inspection };
  inspection.matches = inspection.matches.map(([colIndex, matchKey]) => ({ colIndex, key: matchKey }));
  payload[key] = {
    filename: config.filename,
    inspection,
    base64: fs.readFileSync(path.join(root, "exact-templates", config.source)).toString("base64"),
  };
}

const output = `(function (root, factory) {\n  const templates = factory();\n  if (typeof module === "object" && module.exports) module.exports = templates;\n  else root.DEFAULT_TEMPLATE_FILES = templates;\n})(typeof self !== "undefined" ? self : this, function () {\n  return ${JSON.stringify(payload)};\n});\n`;
fs.writeFileSync(path.join(root, "default-templates.js"), output, "utf8");

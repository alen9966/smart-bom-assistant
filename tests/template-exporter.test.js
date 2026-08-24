const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");
const XLSX = require("xlsx-js-style");
const Core = require("../bom-core.js");
const Exporter = require("../template-exporter.js");
const Definitions = require("../default-templates.js");

const rows = [
  { category: "电阻", partNumber: "R-001", description: "共视接收机主板精密贴片电阻阻值二十二欧姆耐压五十伏封装零四零二", model: "RC02W22R0FT0402B104K500NTEXTRA1234567890ABCDEF", footprint: "0402R", quality: "I", designator: "R1, R2, R3, R4, R5, R6, R7, R8, R9, R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R20", baseQuantity: 2, quantityRaw: "2", quantity: 8, unit: "只", vendor: "风华", source: "外购", drawingNo: "", assemblyName: "主板", remark: "电装", sourceFile: "验证BOM.xlsx", sourceSheet: "BOM", sourceRow: 7, errors: [], warnings: [] },
  { category: "电容", partNumber: "C-001", description: "100nF/50V", model: "0402B104K500NT", footprint: "0402C", quality: "I", designator: "C1", baseQuantity: 4, quantityRaw: "4", quantity: 16, unit: "只", vendor: "国巨", source: "外购", drawingNo: "", assemblyName: "主板", remark: "电装", sourceFile: "验证BOM.xlsx", sourceSheet: "BOM", sourceRow: 8, errors: [], warnings: [] },
];
const metadata = { projectCode: "ZCJY260704", productModel: "共视接收机", projectName: "整机采购", department: "设备与系统", batch: "2608", multiplier: 4, pcbRevision: "ZC7.820.0155-V3", pcbVendor: "嘉立创", sourceFile: "验证BOM.xlsx", mode: "bom" };
const outputDir = path.join(__dirname, "tmp");
fs.mkdirSync(outputDir, { recursive: true });

async function zipText(bytes, name) {
  const zip = await JSZip.loadAsync(bytes);
  const entry = zip.file(name);
  return entry ? entry.async("string") : "";
}

function styleId(xml, address) {
  const cell = xml.match(new RegExp(`<c[^>]*r=["']${address}["'][^>]*>`, "i"));
  const style = cell && cell[0].match(/\bs=["'](\d+)["']/i);
  return style ? style[1] : "";
}

function styleAlignment(stylesXml, id) {
  const body = (stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i) || [])[1] || "";
  const xfs = [...body.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/gi)].map((item) => item[0]);
  const xf = xfs[Number(id)] || "";
  return (xf.match(/<alignment\b[^>]*\/>/i) || [])[0] || "";
}

async function run() {
  const templates = Object.fromEntries(Object.entries(Definitions).map(([key, definition]) => [key, { exactDefinition: definition }]));
  const result = await Exporter.buildFiles(rows, metadata, { selectedOutputs: ["assembly", "picking", "procurement", "subcontract"], skipInvalid: true, purchaseMode: "auto", templates }, Core, XLSX);
  assert.equal(result.files.length, 4, "四个常用公司模板应分别生成四个文件");

  for (const file of result.files) {
    const definition = Definitions[file.key];
    const zip = await JSZip.loadAsync(file.bytes);
    assert.equal(zip.file("xl/calcChain.xml"), null, `${file.key} 不应保留失效计算链`);
    assert.equal(Object.keys(zip.files).some((name) => /^xl\/externalLinks\//i.test(name)), false, `${file.key} 不应保留外部链接`);
    assert.doesNotMatch(await zipText(file.bytes, "[Content_Types].xml"), /calcChain|externalLink/i);
    assert.doesNotMatch(await zipText(file.bytes, "xl/_rels/workbook.xml.rels"), /calcChain|externalLink/i);
    const expectedView = file.key === "assembly" ? "normal" : "pageBreakPreview";
    assert.match(await zipText(file.bytes, definition.inspection.sheetPath), new RegExp(`<sheetView[^>]*view=["']${expectedView}["'][^>]*topLeftCell=["']A1["']`, "i"));

    const originalZip = await JSZip.loadAsync(Buffer.from(definition.base64, "base64"));
    for (const name of ["xl/theme/theme1.xml", "xl/printerSettings/printerSettings1.bin"]) {
      const original = originalZip.file(name);
      const generated = zip.file(name);
      if (original && generated) assert.deepEqual(await generated.async("uint8array"), await original.async("uint8array"), `${file.key} 的 ${name} 必须逐字节保留`);
    }
    // styles.xml：原有 xf 可补 wrapText；末尾可追加左对齐克隆样式
    {
      const originalStyles = originalZip.file("xl/styles.xml");
      if (originalStyles) {
        const original = await originalStyles.async("string");
        const generated = await zipText(file.bytes, "xl/styles.xml");
        const listXfs = (xml) => {
          const body = (xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i) || [])[1] || "";
          return [...body.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/gi)].map((item) => item[0]);
        };
        const normalize = (xml) => xml.replace(/\s*wrapText=["']1["']/gi, "");
        const originalXfs = listXfs(original);
        const generatedXfs = listXfs(generated);
        assert.ok(generatedXfs.length >= originalXfs.length, `${file.key} 不得删除原有样式`);
        originalXfs.forEach((xf, index) => {
          assert.equal(normalize(generatedXfs[index]), normalize(xf), `${file.key} 原有样式 ${index} 除 wrapText 外必须保持一致`);
        });
      }
    }

    const outputPath = path.join(outputDir, `exact-${file.key}.xlsx`);
    fs.writeFileSync(outputPath, file.bytes);
    const workbook = XLSX.readFile(outputPath, { cellStyles: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const outputXml = await zipText(file.bytes, definition.inspection.sheetPath);
    if (file.key === "assembly") {
      assert.equal(String(sheet.B6.v).replace(/\n/g, ""), "共视接收机主板精密贴片电阻阻值二十二欧姆耐压五十伏封装零四零二", "长名称必须完整写入");
      assert.match(String(sheet.B6.v), /\n/, "长名称应按列宽软换行，避免单行裁切");
      assert.equal(String(sheet.C6.v).replace(/\n/g, ""), "RC02W22R0FT0402B104K500NTEXTRA1234567890ABCDEF", "长型号必须完整写入");
      assert.match(String(sheet.C6.v), /\n/, "长型号应按列宽软换行，避免单行裁切");
      assert.equal(sheet.E6.v, 2);
      assert.equal(sheet.F6.v, 8);
      assert.equal(sheet.G6.v, "风华");
      assert.equal(sheet.H6.v, "0402R");
      assert.match(String(sheet.D6.v), /R1[\s\S]*R20/, "长位号必须完整写入，不能截断");
      assert.match(outputXml, /<row\b[^>]*r=["']6["'][^>]*ht=["']([3-9]\d|\d{3,})/, "长文本行应自动加高");
      assert.match(outputXml, /<c[^>]*r=["']F6["'][^>]*>[\s\S]*?<f>E6\*\$G\$3<\/f>[\s\S]*?<v>8<\/v>/i, "领料数应为单机数×本次装配数量的公式");
      assert.match(outputXml, /<c[^>]*r=["']G3["'][^>]*>[\s\S]*?<v>4<\/v>/i, "本次装配数量应写入 G3");
      assert.doesNotMatch(outputXml, /<f>[\s\S]*?(?:SUM|VLOOKUP|INDIRECT)/i, "不应残留模板中的失效公式");
      const stylesXml = await zipText(file.bytes, "xl/styles.xml");
      assert.match(stylesXml, /wrapText=["']1["']/, "装配清单样式应启用自动换行");
      assert.match(styleAlignment(stylesXml, styleId(outputXml, "B6")), /horizontal=["']left["']/i, "名称应左对齐");
      assert.match(styleAlignment(stylesXml, styleId(outputXml, "C6")), /horizontal=["']left["']/i, "型号应左对齐");
      assert.match(styleAlignment(stylesXml, styleId(outputXml, "D6")), /horizontal=["']left["']/i, "位号应左对齐");
      assert.match(outputXml, /<col\b[^>]*min=["']4["'][^>]*width=["'](1\d(?:\.\d+)?)["']/i, "位号列应收窄以便给备注留位");
      assert.match(outputXml, /<col\b[^>]*min=["']9["'][^>]*width=["']([1-9]\d(?:\.\d+)?)["']/i, "备注列应加宽保证可打印");
      assert.match(outputXml, /fitToWidth=["']1["']/i, "装配清单应按页宽适应打印，避免备注列被裁切");
      assert.match(outputXml, /<row\b[^>]*r=["']50["'][^>]*hidden=["']1["']/i, "未使用的空数据行应隐藏，避免打印空白页");
      const workbookXml = await zipText(file.bytes, "xl/workbook.xml");
      assert.match(workbookXml, /_xlnm\.Print_Area[\s\S]*?\$I\$7/i, "装配清单打印区域应截止到实际数据行");
    } else if (file.key === "picking") {
      assert.equal(sheet.C4.v, "电阻");
      assert.equal(String(sheet.C5.v).replace(/\n/g, ""), "RC02W22R0FT0402B104K500NTEXTRA1234567890ABCDEF", "领料单名称型号须完整");
      assert.match(String(sheet.C5.v), /\n/, "领料单长名称型号应软换行");
      assert.equal(sheet.E5.v, 2);
      assert.equal(sheet.G5.v, 8);
      assert.equal(styleId(outputXml, "C4"), "19", "领料单分类行必须使用宋体 9 号加粗样式");
      assert.match(styleAlignment(await zipText(file.bytes, "xl/styles.xml"), styleId(outputXml, "C5")), /horizontal=["']left["']/i, "领料单名称型号应左对齐");
      assert.equal(sheet.K5.v, "电装");
      assert.equal(styleId(outputXml, "K5"), "22", "领料单电装备注必须使用模板红色加粗样式");
      assert.match(outputXml, /<oddFooter>&amp;C第 &amp;P 页，共 &amp;N 页<\/oddFooter>/);
    } else if (file.key === "procurement") {
      assert.equal(sheet.C4.v, "电阻");
      assert.equal(String(sheet.C5.v).replace(/\n/g, ""), "RC02W22R0FT0402B104K500NTEXTRA1234567890ABCDEF", "采购清单名称型号须完整");
      assert.match(String(sheet.C5.v), /\n/, "采购清单长名称型号应软换行");
      assert.equal(sheet.F5.v, 2);
      assert.equal(sheet.H5.v, "是");
      assert.equal(sheet.J5.v, 8);
      assert.equal(sheet.K5.v, 8);
      assert.equal(styleId(outputXml, "C4"), "6", "采购单分类行必须复用模板加粗样式");
      assert.match(styleAlignment(await zipText(file.bytes, "xl/styles.xml"), styleId(outputXml, "C5")), /horizontal=["']left["']/i, "采购清单名称型号应左对齐");
      assert.notEqual(styleId(outputXml, "C5"), "6", "采购单物料行不应误用分类行样式");
      assert.match(await zipText(file.bytes, "xl/styles.xml"), /wrapText=["']1["']/, "采购清单样式应启用自动换行");
      assert.match(outputXml, /<col\b[^>]*min=["']3["'][^>]*width=["']([0-3]?\d(?:\.\d+)?)["']/i, "采购清单名称型号列宽应收窄以适配 A4");
      assert.match(outputXml, /fitToWidth=["']1["']/i, "采购清单应按页宽适应打印，避免备注列被裁切");
      assert.match(outputXml, /<col\b[^>]*min=["']13["'][^>]*width=/i, "采购清单应设置备注列宽度");
      assert.match(outputXml, /<row\b[^>]*r=["']50["'][^>]*hidden=["']1["']/i, "采购清单未使用空行应隐藏，避免打印空白页");
      assert.match(outputXml, /<row\b[^>]*r=["']92["'][^>]*>[\s\S]*?<\/row>/i, "采购单必须保留图三的复选项行");
      assert.match(outputXml, /<row\b[^>]*r=["']93["'][^>]*>[\s\S]*?<\/row>/i, "采购单必须保留图三的签字行");
      assert.doesNotMatch(outputXml, /<row\b[^>]*r=["']92["'][^>]*hidden=["']1["']/i, "表尾行不得被隐藏");
      assert.match(outputXml, /<oddFooter>&amp;C第 &amp;P 页，共 &amp;N 页<\/oddFooter>/);
      assert.equal(sheet.C8.v, "印制板", "采购清单应追加印制板分类行");
      assert.equal(String(sheet.C9.v), "ZC7.820.0155-V3", "印制板名称应为版号");
      assert.equal(sheet.G9.v, "嘉立创", "印制板生产厂家应为页面 PCB 板厂");
      assert.equal(sheet.D9.v, "I", "印制板质量等级应与公司模板一致");
      assert.equal(sheet.F9.v, 1, "印制板单机数量默认为 1");
      assert.equal(sheet.J9.v, 4, "印制板装机总数应等于生产数量");
      assert.equal(sheet.K9.v, 4, "印制板采购总数应等于生产数量");
      assert.equal(styleId(outputXml, "C8"), "6", "印制板分类行必须复用模板加粗样式");
      assert.notEqual(styleId(outputXml, "C9"), "6", "印制板物料行不应误用分类行样式");
    } else {
      assert.equal(sheet.C4.v, "电阻");
      assert.equal(String(sheet.C5.v).replace(/\n/g, ""), "RC02W22R0FT0402B104K500NTEXTRA1234567890ABCDEF", "外协阻容名称型号须完整");
      assert.match(String(sheet.C5.v), /\n/, "外协阻容长名称型号应软换行");
      assert.equal(sheet.F5.v, 2, "外协阻容单机数量必须保持原数量");
      assert.equal(sheet.J5.v, 8, "外协阻容装机总数必须乘生产数量");
      assert.equal(sheet.K5.v, 8, "外协阻容采购总数必须乘生产数量");
      assert.equal(styleId(outputXml, "C4"), "6", "外协阻容分类行必须复用模板加粗样式");
      assert.match(styleAlignment(await zipText(file.bytes, "xl/styles.xml"), styleId(outputXml, "C5")), /horizontal=["']left["']/i, "外协阻容名称型号应左对齐");
      assert.match(await zipText(file.bytes, "xl/styles.xml"), /wrapText=["']1["']/, "外协阻容样式应启用自动换行");
      assert.match(outputXml, /fitToWidth=["']1["']/i, "外协阻容应按页宽适应打印");
      assert.match(outputXml, /<row\b[^>]*r=["']92["'][^>]*>[\s\S]*?<\/row>/i, "外协阻容必须保留复选项行");
      assert.match(outputXml, /<row\b[^>]*r=["']93["'][^>]*>[\s\S]*?<\/row>/i, "外协阻容必须保留签字行");
      assert.match(outputXml, /<oddFooter>&amp;C第 &amp;P 页，共 &amp;N 页<\/oddFooter>/);
    }
    if (file.key !== "assembly") {
      const workbookXml = await zipText(file.bytes, "xl/workbook.xml");
      assert.match(workbookXml, new RegExp(`_xlnm\\.Print_Area[\\s\\S]*?\\$${definition.inspection.printAreaMaxCol}\\$\\d+`, "i"));
      assert.match(workbookXml, /_xlnm\.Print_Titles/i);
    }
  }
  console.log("exact template preservation and Excel compatibility tests passed");
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

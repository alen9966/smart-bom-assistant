const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx-js-style");
const JSZip = require("jszip");

async function run() {
  const directory = process.argv[2];
  const filename = fs.readdirSync(directory).find((name) => name.includes("外协阻容") && name.endsWith(".xlsx"));
  if (!filename) throw new Error("未找到外协阻容输出文件");
  const file = path.join(directory, filename);
  const workbook = XLSX.readFile(file, { cellStyles: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file("xl/worksheets/sheet1.xml").async("string");
  const stylesXml = await zip.file("xl/styles.xml").async("string");
  const cell = (address) => (xml.match(new RegExp(`<c[^>]*r=["']${address}["'][^>]*>[\\s\\S]*?<\\/c>`, "i")) || [])[0] || "";
  const styleId = (address) => (cell(address).match(/\bs=["'](\d+)/) || [])[1] || "";
  const xfs = [...((stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i) || ["", ""])[1]).matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map((match) => match[0]);
  const fonts = [...((stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i) || ["", ""])[1]).matchAll(/<font>([\s\S]*?)<\/font>/gi)].map((match) => match[1]);
  const xf = xfs[Number(styleId("C4"))] || "";
  const fontId = Number((xf.match(/fontId=["'](\d+)/) || [])[1]);
  console.log(JSON.stringify({
    file,
    range: sheet["!ref"],
    firstCategory: sheet.C4.v,
    hasResistorCategory: Array.from({ length: 88 }, (_, index) => sheet[`C${index + 4}`] && sheet[`C${index + 4}`].v).includes("电阻"),
    firstModel: sheet.C5.v,
    baseQuantity: sheet.F5.v,
    totalQuantity: sheet.J5.v,
    purchaseQuantity: sheet.K5.v,
    categoryStyle: styleId("C4"),
    categoryBold: /<b\/?/.test(fonts[fontId] || ""),
    footer92: sheet.A92.v,
    footer93: sheet.A93.v,
  }, null, 2));
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

const root = path.resolve(__dirname, "..");
const targets = ["default-procurement.xlsx", "default-subcontract.xlsx"];

// A4 横向可打印宽度内的列宽（字符宽）。C 列原 88.625 过大，是备注被裁切主因。
const colsXml = `<cols>
<col min="1" max="1" width="4.5" customWidth="1"/>
<col min="2" max="2" width="10" customWidth="1"/>
<col min="3" max="3" width="28" customWidth="1"/>
<col min="4" max="4" width="6.5" customWidth="1"/>
<col min="5" max="5" width="8" customWidth="1"/>
<col min="6" max="6" width="6" customWidth="1"/>
<col min="7" max="7" width="12" style="1" customWidth="1"/>
<col min="8" max="8" width="7" customWidth="1"/>
<col min="9" max="9" width="5.5" customWidth="1"/>
<col min="10" max="10" width="6.5" customWidth="1"/>
<col min="11" max="11" width="6.5" customWidth="1"/>
<col min="12" max="12" width="10" customWidth="1"/>
<col min="13" max="13" width="14" customWidth="1"/>
<col min="14" max="14" width="10" customWidth="1"/>
<col min="15" max="15" width="10" customWidth="1"/>
<col min="16" max="16" width="8" customWidth="1"/>
</cols>`.replace(/\n/g, "");

function patchSheet(xml) {
  let next = xml;
  if (/<cols>[\s\S]*?<\/cols>/i.test(next)) next = next.replace(/<cols>[\s\S]*?<\/cols>/i, colsXml);
  else next = next.replace(/<(?:x:)?sheetData\b/i, `${colsXml}<sheetData`);

  if (/<sheetPr\b/i.test(next)) {
    if (/<pageSetUpPr\b/i.test(next)) {
      next = next.replace(/<pageSetUpPr\b[^>]*\/>/i, '<pageSetUpPr fitToPage="1"/>');
    } else if (/<sheetPr\b[^>]*\/>/i.test(next)) {
      next = next.replace(/<sheetPr\b([^>]*)\/>/i, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
    } else {
      next = next.replace(/<sheetPr\b([^>]*)>/i, '<sheetPr$1><pageSetUpPr fitToPage="1"/>');
    }
  } else {
    next = next.replace(/<worksheet\b[^>]*>/i, (open) => `${open}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`);
  }

  if (/<pageSetup\b/i.test(next)) {
    next = next.replace(/<pageSetup\b([^>]*)\/>/i, (whole, attrs) => {
      let cleaned = attrs
        .replace(/\s+scale=["'][^"']*["']/gi, "")
        .replace(/\s+fitToWidth=["'][^"']*["']/gi, "")
        .replace(/\s+fitToHeight=["'][^"']*["']/gi, "");
      return `<pageSetup${cleaned} fitToWidth="1" fitToHeight="0"/>`;
    });
  }
  return next;
}

(async () => {
  for (const name of targets) {
    const file = path.join(root, "exact-templates", name);
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const sheetPath = "xl/worksheets/sheet1.xml";
    const xml = await zip.file(sheetPath).async("string");
    zip.file(sheetPath, patchSheet(xml));
    const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
    fs.writeFileSync(file, bytes);
    console.log("patched", name, bytes.length);
  }
})();

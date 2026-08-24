const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

const root = path.resolve(__dirname, "..");

function ensureFitToPage(xml) {
  let next = xml;
  if (/<sheetPr\b/i.test(next)) {
    if (/<pageSetUpPr\b/i.test(next)) {
      next = next.replace(/<pageSetUpPr\b[^>]*\/>/i, '<pageSetUpPr fitToPage="1"/>');
    } else if (/<sheetPr\b[^>]*\/>/i.test(next)) {
      next = next.replace(/<(?:x:)?sheetPr\b([^>]*)\/>/i, '<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>');
    } else {
      next = next.replace(/<(?:x:)?sheetPr\b([^>]*)>/i, '<sheetPr$1><pageSetUpPr fitToPage="1"/>');
    }
  } else {
    next = next.replace(/<(?:x:)?worksheet\b[^>]*>/i, (open) => `${open}<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`);
  }
  if (/<pageSetup\b/i.test(next)) {
    next = next.replace(/<(?:x:)?pageSetup\b([^>]*)\/>/i, (whole, attrs) => {
      let cleaned = attrs
        .replace(/\s+scale=["'][^"']*["']/gi, "")
        .replace(/\s+fitToWidth=["'][^"']*["']/gi, "")
        .replace(/\s+fitToHeight=["'][^"']*["']/gi, "");
      return `<pageSetup${cleaned} fitToWidth="1" fitToHeight="0"/>`;
    });
  }
  return next;
}

async function patchAssembly() {
  const file = path.join(root, "exact-templates", "default-assembly.xlsx");
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const sheetPath = "xl/worksheets/sheet1.xml";
  let xml = await zip.file(sheetPath).async("string");
  // A序号 B名称 C型号 D位号 E单机 F领料 G厂家 H封装 I备注
  const cols = `<cols>
<col min="1" max="1" width="4.5" style="3" customWidth="1"/>
<col min="2" max="2" width="12" style="3" customWidth="1"/>
<col min="3" max="3" width="14" style="3" customWidth="1"/>
<col min="4" max="4" width="18" style="3" customWidth="1"/>
<col min="5" max="5" width="6.5" style="3" customWidth="1"/>
<col min="6" max="6" width="6.5" style="4" customWidth="1"/>
<col min="7" max="7" width="8" style="3" customWidth="1"/>
<col min="8" max="8" width="7" style="3" customWidth="1"/>
<col min="9" max="9" width="12" style="3" customWidth="1"/>
</cols>`.replace(/\n/g, "");
  xml = xml.replace(/<cols>[\s\S]*?<\/cols>/i, cols);
  xml = ensureFitToPage(xml);
  zip.file(sheetPath, xml);
  const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  fs.writeFileSync(file, bytes);
  console.log("assembly patched", bytes.length);
}

async function tightenProcurementLike(name) {
  const file = path.join(root, "exact-templates", name);
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const sheetPath = "xl/worksheets/sheet1.xml";
  let xml = await zip.file(sheetPath).async("string");
  // 再压缩名称型号，给备注更多空间；日期/采购员列略收
  const cols = `<cols>
<col min="1" max="1" width="4" customWidth="1"/>
<col min="2" max="2" width="8" customWidth="1"/>
<col min="3" max="3" width="22" customWidth="1"/>
<col min="4" max="4" width="6" customWidth="1"/>
<col min="5" max="5" width="7" customWidth="1"/>
<col min="6" max="6" width="5.5" customWidth="1"/>
<col min="7" max="7" width="10" style="1" customWidth="1"/>
<col min="8" max="8" width="6" customWidth="1"/>
<col min="9" max="9" width="5" customWidth="1"/>
<col min="10" max="10" width="6" customWidth="1"/>
<col min="11" max="11" width="6" customWidth="1"/>
<col min="12" max="12" width="9" customWidth="1"/>
<col min="13" max="13" width="16" customWidth="1"/>
<col min="14" max="14" width="8" customWidth="1"/>
<col min="15" max="15" width="8" customWidth="1"/>
<col min="16" max="16" width="7" customWidth="1"/>
</cols>`.replace(/\n/g, "");
  xml = xml.replace(/<cols>[\s\S]*?<\/cols>/i, cols);
  xml = ensureFitToPage(xml);
  zip.file(sheetPath, xml);
  const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
  fs.writeFileSync(file, bytes);
  console.log(name, "patched", bytes.length);
}

(async () => {
  await patchAssembly();
  await tightenProcurementLike("default-procurement.xlsx");
  await tightenProcurementLike("default-subcontract.xlsx");
})();

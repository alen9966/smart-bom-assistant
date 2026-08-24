const JSZip = require("jszip");
const definitions = require("../default-templates.js");

async function run() {
  for (const [key, definition] of Object.entries(definitions)) {
    const zip = await JSZip.loadAsync(Buffer.from(definition.base64, "base64"));
    console.log(`### ${key}`);
    console.log(Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort().join("\n"));
    for (const name of ["[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", definition.inspection.sheetPath]) {
      const entry = zip.file(name);
      if (!entry) continue;
      const xml = await entry.async("string");
      const findings = xml.match(/<[^>]*(?:external|calcChain|calcPr|sheetView)[^>]*>/gi) || [];
      console.log(`-- ${name} --`);
      console.log(findings.join("\n") || "(none)");
    }
    for (const name of Object.keys(zip.files).filter((entryName) => /\.xml(?:\.rels)?$/i.test(entryName))) {
      const xml = await zip.file(name).async("string");
      const formulaLinks = xml.match(/<f[^>]*>[\s\S]*?<\/f>|<definedName[^>]*>[\s\S]*?<\/definedName>/gi) || [];
      const suspicious = formulaLinks.filter((item) => /\[|http:|https:|file:/i.test(item));
      if (suspicious.length) {
        console.log(`-- suspicious formulas in ${name} --`);
        console.log(suspicious.join("\n"));
      }
    }
    const sheet = await zip.file(definition.inspection.sheetPath).async("string");
    for (const rowNumber of [4, 5, 6, 7, 8, 9, 10, 16, 22, 60, 61, 62, 63, 64, 65]) {
      const row = (sheet.match(new RegExp(`<row[^>]*r=["']${rowNumber}["'][\\s\\S]*?</row>`, "i")) || [""])[0];
      console.log(`row ${rowNumber}: ${(row.match(/<c[^>]*r=["'][A-Z]+\d+["'][^>]*>/gi) || []).join(" ")}`);
    }
  }
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

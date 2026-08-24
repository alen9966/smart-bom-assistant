const JSZip = require("jszip");
const definitions = require("../default-templates.js");

async function run() {
  for (const key of ["picking", "procurement"]) {
    const definition = definitions[key];
    const zip = await JSZip.loadAsync(Buffer.from(definition.base64, "base64"));
    const sheet = await zip.file(definition.inspection.sheetPath).async("string");
    const workbook = await zip.file("xl/workbook.xml").async("string");
    const styles = await zip.file("xl/styles.xml").async("string");
    console.log(`### ${key}`);
    for (const pattern of [
      /<sheetViews>[\s\S]*?<\/sheetViews>/i,
      /<rowBreaks[\s\S]*?<\/rowBreaks>/i,
      /<pageMargins[^>]*\/>/i,
      /<pageSetup[^>]*\/>/i,
      /<headerFooter[\s\S]*?<\/headerFooter>/i,
      /<printOptions[^>]*\/>/i,
      /<definedNames>[\s\S]*?<\/definedNames>/i,
    ]) console.log((sheet.match(pattern) || workbook.match(pattern) || ["(none)"])[0]);
    const rows = [...sheet.matchAll(/<row\b[^>]*\br=["'](\d+)["'][^>]*>[\s\S]*?<\/row>/gi)];
    rows.slice(-7).forEach((match) => console.log(`row ${match[1]} ${match[0].slice(0, 900)}`));
    console.log(`styles fonts/cellxfs: ${(styles.match(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/i) || [""])[0].slice(0, 2600)}\n${(styles.match(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/i) || [""])[0].slice(0, 3600)}`);
  }
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

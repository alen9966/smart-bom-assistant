const JSZip = require("jszip");
const definitions = require("../default-templates.js");

async function run() {
  for (const key of ["picking", "procurement"]) {
    const zip = await JSZip.loadAsync(Buffer.from(definitions[key].base64, "base64"));
    const xml = await zip.file("xl/styles.xml").async("string");
    const fonts = (xml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i) || [])[1] || "";
    const xfs = (xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i) || [])[1] || "";
    console.log(`### ${key}`);
    [...fonts.matchAll(/<font>([\s\S]*?)<\/font>/gi)].forEach((match, index) => console.log(`font ${index}: ${match[1]}`));
    [...xfs.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].forEach((match, index) => console.log(`xf ${index}: ${match[0]}`));
  }
}
run().catch((error) => { console.error(error); process.exitCode = 1; });

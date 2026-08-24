const fs = require("node:fs");
const path = require("node:path");
const ZipUtils = require("../zip-utils.js");

async function run() {
  const sourceDir = path.resolve(process.argv[2] || path.join(__dirname, "..", ".qa-default-templates", "real-v17"));
  const output = path.resolve(process.argv[3] || path.join(__dirname, "..", ".qa-zip-v18", "multi-output.zip"));
  const files = fs.readdirSync(sourceDir).filter((name) => /^ZCJY.*\.xlsx$/i.test(name));
  const entries = files.map((filename) => ({ path: filename, bytes: fs.readFileSync(path.join(sourceDir, filename)) }));
  const packed = await ZipUtils.createZipBytes(entries);
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, packed.bytes);
  const reopened = await ZipUtils.createZipBytes(entries);
  if (Buffer.compare(Buffer.from(packed.bytes), Buffer.from(reopened.bytes)) !== 0) throw new Error("同一批文件重复打包结果不一致");
  console.log(JSON.stringify({ output, bytes: packed.bytes.length, files: packed.paths }, null, 2));
}

run().catch((error) => { console.error(error); process.exitCode = 1; });

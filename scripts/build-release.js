const fs = require("node:fs");
const path = require("node:path");
const JSZip = require("jszip");

const root = path.resolve(__dirname, "..");
const packageJson = require(path.join(root, "package.json"));
const outputName = `智料单_网页版_v${packageJson.version.replace(/\.0$/, "")}.zip`;
const outputPath = path.join(root, outputName);

const releaseFiles = [
  "index.html",
  "styles.css",
  "app.js",
  "bom-core.js",
  "excel-reader.js",
  "default-templates.js",
  "template-exporter.js",
  "zip-utils.js",
  "README.md",
  "CHANGELOG.md",
  "请先看我.txt",
  "启动料单助手.bat",
  "vendor/cpexcel.full.js",
  "vendor/xlsx.bundle.js",
  "vendor/jszip.min.js"
];

function localReferences(html) {
  return [...html.matchAll(/(?:src|href)=["']([^"']+)["']/g)]
    .map((match) => match[1])
    .filter((value) => !/^(?:https?:|data:|#)/i.test(value));
}

async function build() {
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  const referenced = localReferences(html);
  const required = [...new Set([...releaseFiles, ...referenced])];
  const missingOnDisk = required.filter((relative) => !fs.existsSync(path.join(root, relative)));
  if (missingOnDisk.length) throw new Error(`发布失败，磁盘缺少文件：${missingOnDisk.join("、")}`);

  const zip = new JSZip();
  for (const relative of required) {
    zip.file(relative.replaceAll("\\", "/"), fs.readFileSync(path.join(root, relative)));
  }
  const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 }, platform: "DOS" });
  fs.writeFileSync(outputPath, bytes);

  const check = await JSZip.loadAsync(fs.readFileSync(outputPath));
  const missingInZip = required.filter((relative) => !check.file(relative.replaceAll("\\", "/")));
  if (missingInZip.length) throw new Error(`发布失败，ZIP 缺少文件：${missingInZip.join("、")}`);
  console.log(`已生成 ${outputName}（${bytes.length} 字节，${required.length} 个文件，完整性检查通过）`);
}

build().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const app = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const activeDownloadSection = app.slice(app.indexOf("async function downloadOutputFiles"), app.indexOf("function bindEvents"));

assert.match(activeDownloadSection, /for \(const file of list\)/, "多清单必须逐个下载");
assert.match(activeDownloadSection, /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/, "必须以 Excel MIME 下载");
assert.doesNotMatch(activeDownloadSection, /createZipBytes|application\/zip|\.zip["'`]/, "业务下载路径不得再生成 ZIP");
assert.match(activeDownloadSection, /filename: `\$\{prefix\}_\$\{file\.filename\}`/, "多个 BOM 的文件名必须包含 BOM 前缀");
assert.match(app, /window\.showDirectoryPicker/, "网页版应支持通过浏览器选择保存文件夹");
assert.match(activeDownloadSection, /getFileHandle\(filename, \{ create: true \}\)/, "选择文件夹后应直接创建输出文件");
assert.match(activeDownloadSection, /createWritable\(\)/, "选择文件夹后应直接写入输出文件");

console.log("multiple outputs are downloaded as individual Excel files");

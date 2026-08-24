const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const references = [...html.matchAll(/(?:src|href)=["']([^"']+)["']/g)]
  .map((match) => match[1])
  .filter((value) => !/^(?:https?:|data:|#)/i.test(value));

assert.ok(references.length > 0, "index.html 应引用本地资源");
for (const relative of references) {
  assert.ok(fs.existsSync(path.join(root, relative)), `index.html 引用的文件不存在：${relative}`);
}
assert.match(html, /vendor\/xlsx\.bundle\.js/, "Excel 运行库应从 vendor 目录加载");
assert.match(html, /vendor\/jszip\.min\.js/, "ZIP 运行库应从 vendor 目录加载");

console.log(`release assets are complete (${references.length} local references)`);

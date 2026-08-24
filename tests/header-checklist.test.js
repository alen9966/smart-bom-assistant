const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const app = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const section = app.slice(app.indexOf("function renderHeaderChecklist"), app.indexOf("function processRows"));

assert.match(section, /templateColumnStatus/, "表头区域应按实际模板列检查");
assert.match(section, /state\.selectedOutputs/, "表头检查应按当前勾选清单刷新");
assert.match(section, /present: \{ icon: "✓", label: "已识别" \}/, "应显示已识别模板列");
assert.match(section, /derived: \{ icon: "↗", label: "可推导" \}/, "应显示可推导模板列");
assert.match(section, /optional: \{ icon: "○", label: "可留空" \}/, "应显示允许留空模板列");
assert.match(section, /missing: \{ icon: "!", label: "缺少" \}/, "应显示真正缺少的模板列");

console.log("template column comparison renders all selected template fields and statuses");

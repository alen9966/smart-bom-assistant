const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const packageJson = require(path.join(root, "package.json"));
const main = fs.readFileSync(path.join(root, "desktop", "main.js"), "utf8");
const preload = fs.readFileSync(path.join(root, "desktop", "preload.js"), "utf8");
const app = fs.readFileSync(path.join(root, "app.js"), "utf8");

assert.equal(packageJson.main, "desktop/main.js");
assert.equal(packageJson.build.win.target, "portable");
assert.match(main, /contextIsolation: true/);
assert.match(main, /nodeIntegration: false/);
assert.match(main, /sandbox: true/);
assert.match(main, /showOpenDialog/);
assert.match(main, /path\.basename/);
assert.doesNotMatch(main, /payload\.outputDirectory/);
assert.match(preload, /contextBridge\.exposeInMainWorld\("desktopAPI"/);
assert.match(app, /desktopAPI/);

console.log("desktop bridge and packaging configuration tests passed");

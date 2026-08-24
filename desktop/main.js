const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const fs = require("node:fs");
const path = require("node:path");

let mainWindow;

function settingsPath() {
  return path.join(app.getPath("userData"), "settings.json");
}

function loadSettings() {
  try {
    return JSON.parse(fs.readFileSync(settingsPath(), "utf8"));
  } catch (_) {
    return {};
  }
}

function saveSettings(settings) {
  fs.mkdirSync(path.dirname(settingsPath()), { recursive: true });
  fs.writeFileSync(settingsPath(), JSON.stringify(settings, null, 2), "utf8");
}

function defaultOutputDirectory() {
  const configured = loadSettings().outputDirectory;
  return configured && fs.existsSync(configured) ? configured : app.getPath("documents");
}

function safeFilename(filename) {
  const base = path.basename(String(filename || "清单.xlsx")).replace(/[<>:\"/\\|?*\x00-\x1F]/g, "_");
  return base.toLowerCase().endsWith(".xlsx") ? base : `${base}.xlsx`;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1040,
    minHeight: 720,
    backgroundColor: "#f3f5f1",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  });
  mainWindow.loadFile(path.join(__dirname, "..", "index.html"));
}

ipcMain.handle("output-directory:get", () => defaultOutputDirectory());

ipcMain.handle("output-directory:choose", async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "选择 Excel 清单保存位置",
    defaultPath: defaultOutputDirectory(),
    properties: ["openDirectory", "createDirectory"]
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const outputDirectory = path.resolve(result.filePaths[0]);
  saveSettings({ ...loadSettings(), outputDirectory });
  return outputDirectory;
});

ipcMain.handle("output-files:save", async (_event, payload) => {
  // 保存位置只能由主进程的目录选择器和本机设置决定，渲染页面不能
  // 通过 IPC 任意指定系统路径。
  const outputDirectory = path.resolve(defaultOutputDirectory());
  const files = Array.isArray(payload && payload.files) ? payload.files : [];
  if (!files.length) throw new Error("没有可保存的文件");
  fs.mkdirSync(outputDirectory, { recursive: true });
  const saved = [];
  for (const file of files) {
    const filename = safeFilename(file.filename);
    const extension = path.extname(filename);
    const stem = path.basename(filename, extension);
    let destination = path.join(outputDirectory, filename);
    let suffix = 2;
    while (fs.existsSync(destination) || saved.includes(destination)) {
      destination = path.join(outputDirectory, `${stem} (${suffix})${extension}`);
      suffix += 1;
    }
    const bytes = Buffer.from(file.bytes);
    fs.writeFileSync(destination, bytes);
    saved.push(destination);
  }
  saveSettings({ ...loadSettings(), outputDirectory });
  return { outputDirectory, saved };
});

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });

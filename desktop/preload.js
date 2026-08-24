const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopAPI", Object.freeze({
  isDesktop: true,
  getOutputDirectory: () => ipcRenderer.invoke("output-directory:get"),
  chooseOutputDirectory: () => ipcRenderer.invoke("output-directory:choose"),
  saveOutputFiles: (payload) => ipcRenderer.invoke("output-files:save", payload)
}));

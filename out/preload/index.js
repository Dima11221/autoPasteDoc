"use strict";
const electron = require("electron");
electron.contextBridge.exposeInMainWorld("api", {
  selectDocx: () => electron.ipcRenderer.invoke("dialog:selectDocx"),
  selectPhotosFolder: () => electron.ipcRenderer.invoke("dialog:selectPhotosFolder"),
  insertPhotos: (payload) => electron.ipcRenderer.invoke("insert:run", payload),
  showItemInFolder: (filePath) => electron.ipcRenderer.invoke("shell:showItemInFolder", filePath)
});

// src/preload.js
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  saveFileDialog: (defaultPath) => ipcRenderer.invoke('save-file-dialog', defaultPath),
  logMessage: (message) => ipcRenderer.send('log-message', message)
});

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openExcelDialog: () => ipcRenderer.invoke('open-excel-dialog'),
  readExcelFile: (filePath) => ipcRenderer.invoke('read-excel-file', filePath),
});

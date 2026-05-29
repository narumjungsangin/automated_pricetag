const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('node:path');
const fs = require('node:fs');

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    width: 1280,
    height: 900,
    minWidth: 1100,
    minHeight: 760,
    title: '약국 가격표 생성기',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.cjs'),
    },
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
};

ipcMain.handle('open-excel-dialog', async () => {
  const result = await dialog.showOpenDialog({
    title: '엑셀 DB 파일 선택',
    filters: [{ name: '엑셀 파일', extensions: ['xlsx', 'xls', 'csv'] }],
    properties: ['openFile'],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const filePath = result.filePaths[0];
  const data = fs.readFileSync(filePath);
  return { filePath, data: data.toString('base64') };
});

ipcMain.handle('read-excel-file', async (_event, filePath) => {
  try {
    if (!fs.existsSync(filePath)) return null;
    const data = fs.readFileSync(filePath);
    return { filePath, data: data.toString('base64') };
  } catch {
    return null;
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// main.js
const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const log = require('electron-log');
const { autoUpdater } = require('electron-updater');

let mainWindow;
let updateWindow;

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    icon: path.join(__dirname, 'build', '309616865.jpg'),
    webPreferences: {
      preload: path.join(__dirname, 'src', 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true
    }
  });
  mainWindow.loadFile('index.html');
  mainWindow.on('closed', () => { mainWindow = null; });
}

function createUpdateWindow() {
  // Если окно уже существует, ничего не делаем
  if (updateWindow) return;
  updateWindow = new BrowserWindow({
    width: 400,
    height: 300,
    parent: mainWindow,
    modal: true,
    title: 'Обновление приложения',
    resizable: false,
    minimizable: false,
    maximizable: false,
    webPreferences: {
      // Для update окна можно разрешить nodeIntegration
      nodeIntegration: true,
      contextIsolation: false
    }
  });
  updateWindow.loadFile('update.html');
  updateWindow.on('closed', () => { updateWindow = null; });
  log.info('Update window создано.');
}

app.whenReady().then(() => {
  createMainWindow();

  // Ждем загрузки основного окна
  mainWindow.webContents.on('did-finish-load', () => {
    // Запускаем проверку обновлений вручную
    log.info('Основное окно загружено, начинаем проверку обновлений.');
    autoUpdater.checkForUpdates();
  });

  // Настройка логирования для автообновлений
  autoUpdater.logger = log;
  autoUpdater.logger.transports.file.level = 'info';

  // Если обновление доступно, создаем отдельное окно и передаем информацию
  autoUpdater.on('update-available', (info) => {
    log.info('Обновление доступно:', info);
    createUpdateWindow();
    if (updateWindow) {
      updateWindow.webContents.send('update-info', info);
    }
  });

  // Если обновлений нет, отправляем сообщение в основное окно
  autoUpdater.on('update-not-available', (info) => {
    log.info('Обновлений нет:', info);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Обновлений нет.');
    }
  });

  autoUpdater.on('error', (err) => {
    log.error('Ошибка автообновления:', err);
    if (updateWindow) {
      updateWindow.webContents.send('update-error', err && err.toString());
    } else if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Ошибка автообновления.');
    }
    dialog.showErrorBox('Ошибка автообновления', err == null ? "unknown" : (err.stack || err).toString());
  });

  autoUpdater.on('download-progress', (progressObj) => {
    const percent = Math.round(progressObj.percent);
    log.info(`Скачивание: ${percent}%`);
    if (updateWindow) {
      updateWindow.webContents.send('download-progress', percent);
    }
  });

  autoUpdater.on('update-downloaded', (info) => {
    log.info('Обновление загружено:', info);
    if (updateWindow) {
      updateWindow.webContents.send('update-downloaded', info);
    }
  });
});

ipcMain.on('update-action', (event, action) => {
  log.info('Действие обновления:', action);
  if (action === 'update-now') {
    autoUpdater.downloadUpdate();
  } else if (action === 'later') {
    if (updateWindow) {
      updateWindow.close();
    }
  }
});

ipcMain.on('install-update', () => {
  log.info('Установка обновления');
  autoUpdater.quitAndInstall();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Пример обработчика для диалога сохранения Excel-файла (оставляем ваш код)
ipcMain.handle('save-file-dialog', async (event, defaultPath) => {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: 'Сохранить Excel файл',
    defaultPath,
    filters: [
      { name: 'Excel Files', extensions: ['xlsx'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  return canceled ? null : filePath;
});

ipcMain.on('log-message', (event, message) => {
  log.info(message);
});

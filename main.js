// main.js
const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const log = require('electron-log');
const { autoUpdater } = require('electron-updater');

let mainWindow;

function createWindow() {
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
  mainWindow.on('closed', () => mainWindow = null);

  // Запуск проверки обновлений при создании окна
  autoUpdater.checkForUpdatesAndNotify();
}

app.whenReady().then(() => {
  createWindow();

  // Настройка логирования для autoUpdater
  autoUpdater.logger = log;
  autoUpdater.logger.transports.file.level = 'info';

  // Обработка событий автообновления
  autoUpdater.on('checking-for-update', () => {
    log.info('Проверка обновлений...');
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Проверка обновлений...');
    }
  });

  autoUpdater.on('update-available', (info) => {
    log.info('Обновление доступно.', info);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Обновление доступно. Загрузка...');
    }
  });

  autoUpdater.on('update-not-available', (info) => {
    log.info('Обновлений нет.', info);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Обновлений нет.');
    }
  });

  autoUpdater.on('error', (err) => {
    log.error('Ошибка автообновления: ', err);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Ошибка автообновления.');
    }
  });

  autoUpdater.on('download-progress', (progressObj) => {
    let logMessage = `Скачивание: ${Math.round(progressObj.percent)}%`;
    log.info(logMessage);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', logMessage);
    }
  });

  autoUpdater.on('update-downloaded', (info) => {
    log.info('Обновление загружено.', info);
    if (mainWindow) {
      mainWindow.webContents.send('update-message', 'Обновление загружено. Приложение будет перезапущено для установки обновления.');
    }
    // Автоматически перезапустить приложение через 5 секунд
    setTimeout(() => {
      autoUpdater.quitAndInstall();
    }, 5000);
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Обработчик для выбора пути сохранения файла Excel
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

// Обработчик логов
ipcMain.on('log-message', (event, message) => {
  log.info(message);
});

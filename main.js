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

  // Запускаем проверку обновлений сразу после создания окна
  checkForUpdates();
}

// Функция проверки обновлений
function checkForUpdates() {
  // Запускаем проверку обновлений; все события autoUpdater будут обработаны ниже
  autoUpdater.checkForUpdates();
}

// Настраиваем логирование для autoUpdater
autoUpdater.logger = log;
autoUpdater.logger.transports.file.level = 'info';

autoUpdater.on('checking-for-update', () => {
  console.log('Проверка обновлений...');
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Обновления',
    message: 'Проверка обновлений...'
  });
});

autoUpdater.on('update-not-available', () => {
  console.log('Обновлений нет.');
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Обновления',
    message: 'Обновлений нет.'
  });
});

autoUpdater.on('update-available', (info) => {
  console.log('Обновление доступно:', info);
  const updateMessage = `Доступно обновление!\n\nВерсия: ${info.version}\n\nЧто нового:\n${info.releaseNotes}`;
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Доступно обновление',
    message: updateMessage,
    buttons: ['Обновить', 'Отмена'],
    defaultId: 0,
    cancelId: 1
  }).then(result => {
    if (result.response === 0) {
      autoUpdater.downloadUpdate();
    }
  });
});

autoUpdater.on('download-progress', (progressObj) => {
  let percent = Math.round(progressObj.percent);
  console.log(`Скачивание: ${percent}%`);
  mainWindow.setProgressBar(percent / 100);
});

autoUpdater.on('update-downloaded', () => {
  console.log('Обновление скачано.');
  mainWindow.setProgressBar(-1);
  dialog.showMessageBox(mainWindow, {
    type: 'info',
    title: 'Обновление загружено',
    message: 'Обновление загружено. Перезапустить приложение для установки обновления?',
    buttons: ['Перезапустить', 'Позже'],
    defaultId: 0,
    cancelId: 1
  }).then(result => {
    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
});


app.whenReady().then(() => {
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// Пример обработчика для логирования сообщений из рендерера (если требуется)
ipcMain.on('log-message', (event, message) => {
  log.info(message);
});

// Пример обработчика для сохранения файла (оставляем, если используется в вашем приложении)
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

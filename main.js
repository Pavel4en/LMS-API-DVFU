// main.js
const { app, BrowserWindow, dialog, ipcMain } = require('electron');
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
  mainWindow.on('closed', () => { mainWindow = null; });
}

app.whenReady().then(() => {
  createWindow();

  // Настройка логирования для autoUpdater
  autoUpdater.logger = log;
  autoUpdater.logger.transports.file.level = 'info';

  // Проверяем обновления вручную
  autoUpdater.checkForUpdates();

  // Если обновление доступно, уведомляем пользователя
  autoUpdater.on('update-available', (info) => {
    log.info('Обновление доступно: ', info);
    // Можно вывести сообщение через диалог
    dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Обновление доступно',
      message: 'Доступно новое обновление. Загрузка началась...',
      buttons: ['OK']
    });
    // Сообщение можно отправить в renderer, если требуется:
    mainWindow.webContents.send('update-message', 'Обновление доступно. Загрузка...');
  });

  // При отсутствии обновлений
  autoUpdater.on('update-not-available', (info) => {
    log.info('Обновлений нет: ', info);
    mainWindow.webContents.send('update-message', 'Обновлений нет.');
  });

  // Обработка ошибок
  autoUpdater.on('error', (err) => {
    log.error('Ошибка автообновления: ', err);
    mainWindow.webContents.send('update-message', 'Ошибка автообновления.');
  });

  // Следим за прогрессом загрузки
  autoUpdater.on('download-progress', (progressObj) => {
    let percent = Math.round(progressObj.percent);
    log.info(`Скачивание: ${percent}%`);
    mainWindow.webContents.send('update-message', `Скачивание: ${percent}%`);
  });

  // Когда обновление загружено, предложим пользователю установить его
  autoUpdater.on('update-downloaded', (info) => {
    log.info('Обновление загружено: ', info);
    mainWindow.webContents.send('update-message', 'Обновление загружено. Готово к установке.');
    const response = dialog.showMessageBoxSync(mainWindow, {
      type: 'question',
      buttons: ['Установить сейчас', 'Позже'],
      defaultId: 0,
      title: 'Обновление загружено',
      message: 'Обновление загружено. Хотите установить его сейчас?'
    });
    if (response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// Обработчик для выбора пути сохранения файла Excel (остальной ваш код оставляем)
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

// Если хотите отправлять сообщения об обновлении в renderer, добавьте следующий код в renderer.js:
const { ipcRenderer } = require('electron');
ipcRenderer.on('update-message', (event, message) => {
  console.log('Update message:', message);
  const updateDiv = document.getElementById('updateMessage');
  if (updateDiv) {
    updateDiv.textContent = message;
  }
});

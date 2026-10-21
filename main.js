const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 700,
    title: 'Occess - Sistema de Gestión y Filtrado',
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      preload: path.join(__dirname, 'preload.js')
    },
    backgroundColor: '#0f172a',
    autoHideMenuBar: false,
    show: false
  });

  win.loadFile('index.html');

  win.once('ready-to-show', () => {
    win.maximize();
    win.show();
  });

  // Menú superior elegante
  const template = [
    {
      label: 'Archivo',
      submenu: [
        { label: 'Recargar', accelerator: 'CmdOrCtrl+R', role: 'reload' },
        { label: 'Forzar Recarga', accelerator: 'CmdOrCtrl+Shift+R', role: 'forceReload' },
        { type: 'separator' },
        { label: 'Salir', accelerator: 'Alt+F4', role: 'quit' }
      ]
    },
    {
      label: 'Vista',
      submenu: [
        { label: 'Pantalla Completa', accelerator: 'F11', role: 'togglefullscreen' },
        { label: 'Aumentar Zoom', accelerator: 'CmdOrCtrl+Plus', role: 'zoomIn' },
        { label: 'Disminuir Zoom', accelerator: 'CmdOrCtrl+-', role: 'zoomOut' },
        { label: 'Restablecer Zoom', accelerator: 'CmdOrCtrl+0', role: 'resetZoom' },
        { type: 'separator' },
        { label: 'Herramientas de Desarrollador', accelerator: 'F12', role: 'toggleDevTools' }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

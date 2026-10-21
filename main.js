const { app, BrowserWindow, Menu, dialog, shell } = require('electron');
const path = require('path');
const { autoUpdater } = require('electron-updater');

// Configuración de auto-actualizaciones
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;

let isManualCheck = false;

function setupAutoUpdater(win) {
  autoUpdater.on('checking-for-update', () => {
    console.log('Comprobando actualizaciones en GitHub...');
  });

  autoUpdater.on('update-available', (info) => {
    dialog.showMessageBox(win, {
      type: 'question',
      title: 'Actualización disponible - Occess',
      message: `¡Hay una nueva versión de Occess disponible!`,
      detail: `Versión instalada: v${app.getVersion()}\nNueva versión disponible: v${info.version}\n\n¿Deseas descargarla e instalarla ahora?`,
      buttons: ['Descargar y Actualizar', 'Más tarde'],
      defaultId: 0,
      cancelId: 1
    }).then((result) => {
      if (result.response === 0) {
        autoUpdater.downloadUpdate();
        dialog.showMessageBox(win, {
          type: 'info',
          title: 'Descargando actualización',
          message: 'La actualización se está descargando en segundo plano.',
          detail: 'Recibirás una notificación cuando termine para reiniciar la aplicación.'
        });
      }
    });
  });

  autoUpdater.on('update-not-available', () => {
    if (isManualCheck) {
      dialog.showMessageBox(win, {
        type: 'info',
        title: 'Occess al día',
        message: '¡Tienes la versión más reciente!',
        detail: `Actualmente estás utilizando la versión v${app.getVersion()}.`
      });
      isManualCheck = false;
    }
  });

  autoUpdater.on('update-downloaded', (info) => {
    dialog.showMessageBox(win, {
      type: 'info',
      title: 'Actualización lista para instalar',
      message: `La versión v${info.version} está lista.`,
      detail: 'La aplicación debe reiniciarse para aplicar los cambios. ¿Deseas reiniciar ahora?',
      buttons: ['Reiniciar ahora', 'Más tarde'],
      defaultId: 0
    }).then((result) => {
      if (result.response === 0) {
        autoUpdater.quitAndInstall();
      }
    });
  });

  autoUpdater.on('error', (err) => {
    console.error('Error en autoUpdater:', err);
    if (isManualCheck) {
      dialog.showMessageBox(win, {
        type: 'warning',
        title: 'Aviso de actualización',
        message: 'No se pudo comprobar actualizaciones en este momento.',
        detail: 'Verifica tu conexión a internet o intenta más tarde.'
      });
      isManualCheck = false;
    }
  });
}

function createWindow() {
  const currentVer = app.getVersion();
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1024,
    minHeight: 700,
    title: `Occess v${currentVer} - Sistema de Gestión y Filtrado`,
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
    win.setTitle(`Occess v${currentVer} - Sistema de Gestión y Filtrado`);
    win.maximize();
    win.show();
  });

  // Configurar listeners de auto-update
  setupAutoUpdater(win);

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
    },
    {
      label: 'Ayuda',
      submenu: [
        {
          label: 'Buscar actualizaciones...',
          click: () => {
            isManualCheck = true;
            autoUpdater.checkForUpdates().catch((err) => {
              dialog.showMessageBox(win, {
                type: 'warning',
                title: 'Aviso',
                message: 'No se pudo conectar con el servidor de actualizaciones.',
                detail: err.message
              });
              isManualCheck = false;
            });
          }
        },
        { type: 'separator' },
        {
          label: 'Acerca de Occess',
          click: () => {
            dialog.showMessageBox(win, {
              type: 'info',
              title: 'Acerca de Occess',
              message: `Occess - Sistema de Gestión y Filtrado\nVersión actual: v${app.getVersion()}`,
              detail: 'Sistema optimizado de filtrado y procesamiento rápido de servicios.\n\n© 2026 Occess',
              buttons: ['Aceptar']
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);

  // Comprobar actualizaciones automáticamente 4 segundos tras abrir si está empaquetado
  setTimeout(() => {
    if (app.isPackaged) {
      autoUpdater.checkForUpdates().catch((err) => {
        console.log('Verificación automática inicial:', err.message);
      });
    }
  }, 4000);
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

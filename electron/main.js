const { app, BrowserWindow, Tray, Menu, nativeImage, ipcMain, Notification, screen } = require('electron');
const path = require('path');
const { spawn, execSync } = require('child_process');

const APP_USER_MODEL_ID = 'com.traceability.suite';
const TOAST_ACTIVATOR_CLSID = '{7B8F4B1C-2D84-4A5D-9D3B-0F0B21B83E71}';

// Set Windows app identity as early as possible so toast notifications and
// relaunches stay bound to the packaged app.
app.setName('Project Traceability Suite');
app.setAppUserModelId(APP_USER_MODEL_ID);
if (process.platform === 'win32') {
  app.setToastActivatorCLSID(TOAST_ACTIVATOR_CLSID);
  
  // Establish system trust for the self-signed certificate if not already present
  try {
    const certPath = path.join(__dirname, 'Meet_Mistry_cert.pfx');
    execSync(`powershell -ExecutionPolicy Bypass -Command "if(-not (Get-ChildItem -Path Cert:\\LocalMachine\\Root | Where-Object { $_.Subject -like '*CN=Meet Mistry*' })) { certutil -addstore -f Root '${certPath}'; certutil -addstore -f TrustedPublisher '${certPath}' }"`);
    console.log("System trust check completed.");
  } catch (err) {
    console.log("Certificate trust injection skipped (Admin rights may be required).");
  }
}

// --- Single Instance Protection ---
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
  process.exit(0);
}

const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged;
const DEV_SERVER_URL = 'http://localhost:5173'; // Vite default

const getAppPaths = () => {
  const iconExt = process.platform === 'win32' ? 'ico' : 'png';
  if (app.isPackaged) {
    return {
      backend: path.join(__dirname, 'server', 'server.js'),
      renderer: path.join(__dirname, 'dist', 'index.html'),
      icon: path.join(__dirname, `icon.${iconExt}`),
      splash: path.join(__dirname, 'splash.html')
    };
  }
  return {
    backend: path.join(__dirname, '../backend/server.js'),
    renderer: path.join(__dirname, '../client/dist/index.html'),
    icon: path.join(__dirname, `icon.${iconExt}`),
    splash: path.join(__dirname, 'splash.html')
  };
};

const paths = getAppPaths();

let mainWindow;
let splashWindow;
let tray;
let trayMenu;
let latestTrayStats = null;
let backendProcess;
let splashStartedAt;

const MIN_SPLASH_TIME = 3000; // 3 seconds

function bringMainWindowToFront() {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return false;
  }

  // Ensure window is in its original size/state if minimized
  if (mainWindow.isMinimized()) {
    mainWindow.restore();
  }

  // --- Splash Gate ---
  // If splash is still active, do NOT show the main window. 
  // This prevents background data-loading alerts from popping the window up prematurely.
  const elapsed = Date.now() - (splashStartedAt || 0);
  if (elapsed < MIN_SPLASH_TIME) {
    console.log(`[Splash Gate] bringMainWindowToFront blocked. Remaining: ${MIN_SPLASH_TIME - elapsed}ms`);
    return false;
  }

  // Make it visible if it was hidden in tray
  if (!mainWindow.isVisible()) {
    mainWindow.show();
  }

  // Force to front across all other windows
  mainWindow.setAlwaysOnTop(true);
  mainWindow.setVisibleOnAllWorkspaces(true);
  mainWindow.focus();
  
  // Return to normal layering after a tiny delay so it doesn't block other apps forever
  setTimeout(() => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.setAlwaysOnTop(false);
      mainWindow.setVisibleOnAllWorkspaces(false);
    }
  }, 200);

  return true;
}

function setTrayStats(stats) {
  latestTrayStats = stats;

  if (!trayMenu) {
    return;
  }

  const pendingItem = trayMenu.getMenuItemById('pending-stat');
  const completedItem = trayMenu.getMenuItemById('completed-stat');
  const overdueItem = trayMenu.getMenuItemById('overdue-stat');

  if (pendingItem) pendingItem.label = `Pending: ${stats.pending ?? 0}`;
  if (completedItem) completedItem.label = `Completed: ${stats.completed ?? 0}`;
  if (overdueItem) overdueItem.label = `Overdue: ${stats.overdue ?? 0}`;
}

function startBackend() {
  const dataPath = app.isPackaged 
    ? path.join(app.getPath('userData'), 'data') 
    : path.join(__dirname, '../backend');

  try {
    const fs = require('fs');
    if (!fs.existsSync(dataPath)) {
      fs.mkdirSync(dataPath, { recursive: true });
    }
  } catch (err) {
    console.error('Failed to create data directory:', err);
  }

  // Use the internal electron binary as node to run our server script
  backendProcess = spawn(process.execPath, [paths.backend, `--base-dir=${dataPath}`], {
    stdio: 'inherit',
    env: { 
      ...process.env, 
      NODE_ENV: 'production',
      ELECTRON_RUN_AS_NODE: '1'
    }
  });
}

async function loadRenderer() {
  if (isDev) {
    const maxAttempts = 20;

    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        await mainWindow.loadURL(DEV_SERVER_URL);
        return;
      } catch (error) {
        if (attempt === maxAttempts) {
          throw error;
        }

        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    }

    return;
  }

  await mainWindow.loadFile(paths.renderer);
}

function createSplash() {
  splashWindow = new BrowserWindow({
    width: 500,
    height: 400,
    transparent: true,
    frame: false,
    alwaysOnTop: true,
    show: true,
    backgroundColor: '#0f172a',
    icon: paths.icon,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  splashStartedAt = Date.now();

  splashWindow.loadFile(paths.splash);
  splashWindow.center();
}

async function createWindow() {
  const primaryDisplay = screen.getPrimaryDisplay();
  const { width: workAreaWidth, height: workAreaHeight } = primaryDisplay.workAreaSize;

  const minWidth = 1180;
  const minHeight = 640;

  // Compute a safe initial windowed size (40px smaller than work area)
  const initialWidth = Math.max(minWidth, Math.min(1400, workAreaWidth - 40));
  const initialHeight = Math.max(minHeight, Math.min(900, workAreaHeight - 40));

  const appIconNative = nativeImage.createFromPath(paths.icon);

  mainWindow = new BrowserWindow({
    width: initialWidth,
    height: initialHeight,
    minWidth: minWidth,
    minHeight: minHeight,
    frame: false,
    show: false,
    icon: appIconNative,
    backgroundColor: '#0f172a',
    maximizable: true,
    resizable: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      backgroundThrottling: false
    },
    title: "Project Traceability Suite"
  });

  if (process.platform === 'win32' && !appIconNative.isEmpty()) {
    mainWindow.setIcon(appIconNative);
  }

  mainWindow.on('maximize', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('window-maximized-change', true);
    }
  });

  mainWindow.on('unmaximize', () => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('window-maximized-change', false);
    }
  });

  mainWindow.setFullScreenable(false);
  mainWindow.setMenu(null);

  // Setup ready-to-show BEFORE loading to catch the event early
  mainWindow.once('ready-to-show', () => {
    const showApp = () => {
      if (splashWindow && !splashWindow.isDestroyed()) {
        splashWindow.close();
      }
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.maximize();
        bringMainWindowToFront();
      }
    };

    // Minimum splash duration: Respect MIN_SPLASH_TIME for a premium feel
    const elapsed = Date.now() - (splashStartedAt || Date.now());
    const remainingTime = Math.max(0, MIN_SPLASH_TIME - elapsed);

    setTimeout(showApp, remainingTime); 
  });

  await loadRenderer();

  // Safety fallback in case ready-to-show never fires
  setTimeout(() => {
    if (splashWindow && !splashWindow.isDestroyed()) {
      console.log("ready-to-show timed out, forcing transition...");
      if (splashWindow && !splashWindow.isDestroyed()) splashWindow.close();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.maximize();
        bringMainWindowToFront();
      }
    }
  }, 10000);

  mainWindow.on('close', (event) => {
    if (!app.isQuitting) {
      event.preventDefault();
      mainWindow.hide();
    }
    return false;
  });
}

function createTray() {
  const trayIcon = nativeImage.createFromPath(paths.icon);
  tray = new Tray(trayIcon);

  trayMenu = Menu.buildFromTemplate([
    { label: 'Project Traceability System', enabled: false },
    { type: 'separator' },
    { label: 'Open App', click: () => bringMainWindowToFront() },
    { 
      label: 'Quick Stats', 
      submenu: [
        { label: 'Pending: 0', id: 'pending-stat', enabled: false },
        { label: 'Completed: 0', id: 'completed-stat', enabled: false },
        { label: 'Overdue: 0', id: 'overdue-stat', enabled: false }
      ]
    },
    { type: 'separator' },
    { label: 'Quit', click: () => {
      app.isQuitting = true;
      app.quit();
    }}
  ]);

  tray.setToolTip('Project Traceability System');
  tray.setContextMenu(trayMenu);
  if (latestTrayStats) {
    setTrayStats(latestTrayStats);
  }
}

app.whenReady().then(async () => {
  createSplash();
  
  if (!isDev) {
    startBackend();
  }

  await createWindow();
  createTray();
});

app.on('second-instance', async () => {
  // Windows toast clicks and shortcut activations should restore the existing
  // instance, not spawn a fresh Electron runtime.
  if (!bringMainWindowToFront() && BrowserWindow.getAllWindows().length === 0) {
    await createWindow();
  }
});

app.on('activate', async () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    await createWindow();
  } else if (mainWindow) {
    bringMainWindowToFront();
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    // Keep it alive in tray
  }
});

app.on('quit', () => {
  if (backendProcess) backendProcess.kill();
});

// IPC communication
ipcMain.on('notify', (event, { title, body }) => {
  const notification = new Notification({ 
    title, 
    body,
    icon: paths.icon,
    silent: false 
  });

  notification.on('click', () => {
    bringMainWindowToFront();
  });

  // Windows: Native system notifications
  notification.show();
});

ipcMain.on('update-stats', (_event, stats) => {
  if (!stats || typeof stats !== 'object') {
    return;
  }
  setTrayStats({
    pending: Number(stats.pending) || 0,
    completed: Number(stats.completed) || 0,
    overdue: Number(stats.overdue) || 0
  });
});

ipcMain.on('window-control', (_event, action) => {
  if (!mainWindow) return;
  if (action === 'minimize') {
    mainWindow.minimize();
  }
  if (action === 'close') {
    mainWindow.hide();
  }
  if (action === 'maximize') {
    if (mainWindow.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow.maximize();
    }
  }
  if (action === 'quit') {
    app.isQuitting = true;
    app.quit();
  }
});

ipcMain.on('bring-to-front', () => {
  bringMainWindowToFront();
});

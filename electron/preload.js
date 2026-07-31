const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  notify: (title, body) => ipcRenderer.send('notify', { title, body }),
  onUpdateStats: (callback) => ipcRenderer.on('update-stats', (event, stats) => callback(stats)),
  updateStats: (stats) => ipcRenderer.send('update-stats', stats),
  windowControl: (action) => ipcRenderer.send('window-control', action),
  bringToFront: () => ipcRenderer.send('bring-to-front'),
  onMaximizedStateChange: (callback) => ipcRenderer.on('window-maximized-change', (event, isMaximized) => callback(isMaximized))
});

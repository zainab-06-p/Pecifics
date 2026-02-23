const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods to renderer
contextBridge.exposeInMainWorld('electronAPI', {
    // Settings
    getSettings: () => ipcRenderer.invoke('get-settings'),
    saveSettings: (settings) => ipcRenderer.invoke('save-settings', settings),
    
    // Screenshot
    takeScreenshot: () => ipcRenderer.invoke('take-screenshot'),
    onScreenshotCaptured: (callback) => {
        ipcRenderer.on('screenshot-captured', (event, data) => callback(data));
    },
    
    // Screen info
    getScreenInfo: () => ipcRenderer.invoke('get-screen-info'),
    
    // Window controls
    minimizeWindow: () => ipcRenderer.send('minimize-window'),
    closeWindow: () => ipcRenderer.send('close-window'),
    toggleAlwaysOnTop: (value) => ipcRenderer.send('toggle-always-on-top', value),
    
    // Capture control
    startCapture: () => ipcRenderer.send('start-capture'),
    stopCapture: () => ipcRenderer.send('stop-capture'),
    
    // Action execution (forward to action-executor)
    executeAction: (action) => ipcRenderer.invoke('execute-action', action),
    
    // Stop execution
    stopExecution: () => ipcRenderer.invoke('stop-execution'),
    resetStopFlag: () => ipcRenderer.invoke('reset-stop-flag'),
    checkStopFlag: () => ipcRenderer.invoke('check-stop-flag'),
    
    // Events
    onShowSettings: (callback) => {
        ipcRenderer.on('show-settings', () => callback());
    }
});

// Also expose versions
contextBridge.exposeInMainWorld('versions', {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron
});

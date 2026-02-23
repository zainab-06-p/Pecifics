const { app, BrowserWindow, ipcMain, Tray, Menu, globalShortcut, screen, nativeImage } = require('electron');
const path = require('path');
const Store = require('electron-store');
const screenshot = require('screenshot-desktop');
const sharp = require('sharp');
const actionExecutor = require('./modules/action-executor');

// Initialize store for settings
const store = new Store({
    defaults: {
        colabUrl: 'http://localhost:8001',
        screenshotInterval: 1000, // ms
        screenshotQuality: 80,
        autoScreenshot: true,
        hotkey: 'CommandOrControl+Shift+J'
    }
});

let mainWindow = null;
let tray = null;
let screenshotInterval = null;
let isCapturing = false;

// Create main window
function createWindow() {
    const { width, height } = screen.getPrimaryDisplay().workAreaSize;
    
    mainWindow = new BrowserWindow({
        width: 450,
        height: 700,
        x: width - 470,
        y: height - 720,
        frame: false,
        transparent: true,
        alwaysOnTop: true,
        resizable: true,
        skipTaskbar: false,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        icon: path.join(__dirname, '../assets/icon.png')
    });

    mainWindow.loadFile(path.join(__dirname, 'renderer/index.html'));
    
    // Hide on close, don't quit
    mainWindow.on('close', (event) => {
        if (!app.isQuitting) {
            event.preventDefault();
            mainWindow.hide();
        }
    });

    // DevTools in dev mode
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools({ mode: 'detach' });
    }
}

// Create system tray
function createTray() {
    const iconPath = path.join(__dirname, '../assets/icon.png');
    
    // Create a simple icon if file doesn't exist
    let trayIcon;
    try {
        trayIcon = nativeImage.createFromPath(iconPath);
        if (trayIcon.isEmpty()) {
            trayIcon = nativeImage.createEmpty();
        }
    } catch {
        trayIcon = nativeImage.createEmpty();
    }
    
    tray = new Tray(trayIcon.resize({ width: 16, height: 16 }));
    
    const contextMenu = Menu.buildFromTemplate([
        { 
            label: 'Show Pecifics', 
            click: () => {
                mainWindow.show();
                mainWindow.focus();
            }
        },
        { 
            label: 'Settings', 
            click: () => {
                mainWindow.show();
                mainWindow.webContents.send('show-settings');
            }
        },
        { type: 'separator' },
        { 
            label: 'Toggle Screenshot Capture',
            type: 'checkbox',
            checked: store.get('autoScreenshot'),
            click: (menuItem) => {
                store.set('autoScreenshot', menuItem.checked);
                if (menuItem.checked) {
                    startScreenshotCapture();
                } else {
                    stopScreenshotCapture();
                }
            }
        },
        { type: 'separator' },
        { 
            label: 'Quit Pecifics', 
            click: () => {
                app.isQuitting = true;
                app.quit();
            }
        }
    ]);
    
    tray.setToolTip('Pecifics - AI Desktop Assistant');
    tray.setContextMenu(contextMenu);
    
    tray.on('click', () => {
        if (mainWindow.isVisible()) {
            mainWindow.hide();
        } else {
            mainWindow.show();
            mainWindow.focus();
        }
    });
}

// Take screenshot and return as base64
async function takeScreenshot() {
    try {
        isCapturing = true;
        
        // Capture screenshot
        const imgBuffer = await screenshot({ format: 'png' });
        
        // Get screen dimensions
        const { width, height } = screen.getPrimaryDisplay().size;
        
        // Resize and compress for faster transfer
        const quality = store.get('screenshotQuality');
        const resizedBuffer = await sharp(imgBuffer)
            .resize(Math.floor(width / 2), Math.floor(height / 2)) // Half resolution
            .jpeg({ quality: quality })
            .toBuffer();
        
        const base64 = resizedBuffer.toString('base64');
        
        isCapturing = false;
        
        return {
            screenshot: base64,
            width: width,
            height: height,
            timestamp: Date.now()
        };
    } catch (error) {
        isCapturing = false;
        console.error('Screenshot error:', error);
        return null;
    }
}

// Start continuous screenshot capture
function startScreenshotCapture() {
    if (screenshotInterval) {
        clearInterval(screenshotInterval);
    }
    
    const interval = store.get('screenshotInterval');
    
    screenshotInterval = setInterval(async () => {
        if (!isCapturing && mainWindow && !mainWindow.isDestroyed()) {
            const screenshotData = await takeScreenshot();
            if (screenshotData) {
                mainWindow.webContents.send('screenshot-captured', screenshotData);
            }
        }
    }, interval);
    
    console.log(`Screenshot capture started (interval: ${interval}ms)`);
}

// Stop screenshot capture
function stopScreenshotCapture() {
    if (screenshotInterval) {
        clearInterval(screenshotInterval);
        screenshotInterval = null;
    }
    console.log('Screenshot capture stopped');
}

// Register global hotkey
function registerHotkey() {
    const hotkey = store.get('hotkey');
    
    globalShortcut.unregisterAll();
    
    const success = globalShortcut.register(hotkey, () => {
        if (mainWindow.isVisible()) {
            mainWindow.hide();
        } else {
            mainWindow.show();
            mainWindow.focus();
        }
    });
    
    if (success) {
        console.log(`Hotkey registered: ${hotkey}`);
    } else {
        console.error(`Failed to register hotkey: ${hotkey}`);
    }
}

// App ready
app.whenReady().then(() => {
    createWindow();
    createTray();
    registerHotkey();
    
    if (store.get('autoScreenshot')) {
        startScreenshotCapture();
    }
    
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

// Quit when all windows closed (except on macOS)
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// Cleanup on quit
app.on('will-quit', () => {
    globalShortcut.unregisterAll();
    stopScreenshotCapture();
});

// ============== IPC HANDLERS ==============

// Get settings
ipcMain.handle('get-settings', () => {
    return {
        colabUrl: store.get('colabUrl'),
        screenshotInterval: store.get('screenshotInterval'),
        screenshotQuality: store.get('screenshotQuality'),
        autoScreenshot: store.get('autoScreenshot'),
        hotkey: store.get('hotkey')
    };
});

// Save settings
ipcMain.handle('save-settings', (event, settings) => {
    if (settings.colabUrl !== undefined) store.set('colabUrl', settings.colabUrl);
    if (settings.screenshotInterval !== undefined) store.set('screenshotInterval', settings.screenshotInterval);
    if (settings.screenshotQuality !== undefined) store.set('screenshotQuality', settings.screenshotQuality);
    if (settings.autoScreenshot !== undefined) {
        store.set('autoScreenshot', settings.autoScreenshot);
        if (settings.autoScreenshot) {
            startScreenshotCapture();
        } else {
            stopScreenshotCapture();
        }
    }
    if (settings.hotkey !== undefined) {
        store.set('hotkey', settings.hotkey);
        registerHotkey();
    }
    return true;
});

// Take single screenshot
ipcMain.handle('take-screenshot', async () => {
    return await takeScreenshot();
});

// Get screen info
ipcMain.handle('get-screen-info', () => {
    const primaryDisplay = screen.getPrimaryDisplay();
    return {
        width: primaryDisplay.size.width,
        height: primaryDisplay.size.height,
        scaleFactor: primaryDisplay.scaleFactor
    };
});

// Window controls
ipcMain.on('minimize-window', () => {
    mainWindow.minimize();
});

ipcMain.on('close-window', () => {
    mainWindow.hide();
});

ipcMain.on('toggle-always-on-top', (event, value) => {
    mainWindow.setAlwaysOnTop(value);
});

// Start/stop screenshot capture
ipcMain.on('start-capture', () => {
    startScreenshotCapture();
});

ipcMain.on('stop-capture', () => {
    stopScreenshotCapture();
});

// Action execution
ipcMain.handle('execute-action', async (event, { action, params }) => {
    try {
        const result = await actionExecutor.execute(action, params);
        return result;
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Stop execution
ipcMain.handle('stop-execution', async () => {
    try {
        const result = actionExecutor.stopExecution();
        return result;
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Reset stop flag
ipcMain.handle('reset-stop-flag', async () => {
    try {
        actionExecutor.resetStopFlag();
        return { success: true };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// Check stop flag
ipcMain.handle('check-stop-flag', async () => {
    try {
        return actionExecutor.isStopped();
    } catch (error) {
        return false;
    }
});

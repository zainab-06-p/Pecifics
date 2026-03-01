const { app, BrowserWindow, ipcMain, Tray, Menu, globalShortcut, screen, nativeImage } = require('electron');
const path = require('path');
const Store = require('electron-store');
const screenshot = require('screenshot-desktop');
const sharp = require('sharp');
const actionExecutor = require('./modules/action-executor');

// Initialize store for settings
const store = new Store({
    defaults: {
        colabUrl: 'http://localhost:8000',
        cogagentUrl: '',
        screenshotInterval: 1000, // ms
        screenshotQuality: 80,
        autoScreenshot: true,
        alwaysOnTop: true,
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
        transparent: false,
        backgroundColor: '#0a0a0f',
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
        
        // Capture screenshot — try PNG first, fall back to JPG
        let imgBuffer;
        try {
            imgBuffer = await screenshot({ format: 'png' });
        } catch {
            imgBuffer = await screenshot({ format: 'jpg' });
        }
        
        if (!imgBuffer || imgBuffer.length < 100) {
            isCapturing = false;
            return null;
        }
        
        // Get screen dimensions
        const { width, height } = screen.getPrimaryDisplay().size;
        const quality = store.get('screenshotQuality');

        // Use sharp with failOn:'none' — tolerates slightly corrupt input
        let resizedBuffer;
        try {
            resizedBuffer = await sharp(imgBuffer, { failOn: 'none' })
                .resize(Math.floor(width / 2), Math.floor(height / 2))
                .jpeg({ quality })
                .toBuffer();
        } catch {
            // If PNG decoding fails, try re-capturing as JPEG directly
            try {
                imgBuffer = await screenshot({ format: 'jpg' });
                resizedBuffer = await sharp(imgBuffer, { failOn: 'none' })
                    .resize(Math.floor(width / 2), Math.floor(height / 2))
                    .jpeg({ quality })
                    .toBuffer();
            } catch (e2) {
                isCapturing = false;
                if (!takeScreenshot._lastErr || Date.now() - takeScreenshot._lastErr > 60000) {
                    console.error('Screenshot error:', e2.message);
                    takeScreenshot._lastErr = Date.now();
                }
                return null;
            }
        }
        
        isCapturing = false;
        return {
            screenshot: resizedBuffer.toString('base64'),
            width,
            height,
            timestamp: Date.now()
        };
    } catch (error) {
        isCapturing = false;
        if (!takeScreenshot._lastErr || Date.now() - takeScreenshot._lastErr > 60000) {
            console.error('Screenshot error:', error.message);
            takeScreenshot._lastErr = Date.now();
        }
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
        cogagentUrl: store.get('cogagentUrl'),
        screenshotInterval: store.get('screenshotInterval'),
        screenshotQuality: store.get('screenshotQuality'),
        autoScreenshot: store.get('autoScreenshot'),
        alwaysOnTop: store.get('alwaysOnTop'),
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
    if (settings.cogagentUrl !== undefined) store.set('cogagentUrl', settings.cogagentUrl);
    if (settings.alwaysOnTop !== undefined) {
        store.set('alwaysOnTop', settings.alwaysOnTop);
        if (mainWindow) mainWindow.setAlwaysOnTop(settings.alwaysOnTop);
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

// Take HIGH-RES screenshot for vision agent (no downscaling — Gemini needs accuracy)
ipcMain.handle('take-screenshot-hires', async () => {
    try {
        let imgBuffer;
        try { imgBuffer = await screenshot({ format: 'png' }); }
        catch { imgBuffer = await screenshot({ format: 'jpg' }); }
        if (!imgBuffer || imgBuffer.length < 100) return null;

        const display = screen.getPrimaryDisplay();
        const { width, height } = display.size;
        const scale = display.scaleFactor || 1;

        // Resize to FULL logical resolution (not half) — much better for coordinate accuracy
        // Physical capture may be width*scale, so we resize to exactly width×height logical pixels
        const resized = await sharp(imgBuffer, { failOn: 'none' })
            .resize(width, height)
            .jpeg({ quality: 85 })
            .toBuffer();

        return {
            screenshot: resized.toString('base64'),
            width,       // logical screen width
            height,      // logical screen height
            scaleFactor: scale,
            timestamp: Date.now()
        };
    } catch (e) {
        console.error('HiRes screenshot error:', e.message);
        return null;
    }
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

ipcMain.handle('get-user-home', () => {
    return require('os').homedir();
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

// ============== COGAGENT DIRECT CONNECTION ==============
// Bypasses the langchain backend for vision tasks — calls CogAgent Kaggle directly
// from the main process (avoids CORS issues, custom timeout for ~31s inference)

const axios = require('axios');

// Check CogAgent health
ipcMain.handle('cogagent-health', async () => {
    const url = store.get('cogagentUrl');
    if (!url) return { ok: false, error: 'No CogAgent URL configured' };
    try {
        const resp = await axios.get(`${url.replace(/\/+$/, '')}/health`, { timeout: 10000 });
        return { ok: true, data: resp.data };
    } catch (error) {
        return { ok: false, error: error.message };
    }
});

// Direct vision action — POST screenshot+goal to CogAgent and return the action
ipcMain.handle('cogagent-vision-act', async (event, payload) => {
    const url = store.get('cogagentUrl');
    if (!url) return { action: 'fail', description: 'No CogAgent URL configured. Set it in Settings.' };
    try {
        const resp = await axios.post(`${url.replace(/\/+$/, '')}/vision_act`, {
            screenshot:    payload.screenshot,
            goal:          payload.goal,
            step_history:  payload.step_history || [],
            screen_width:  payload.screen_width || 1920,
            screen_height: payload.screen_height || 1080,
        }, {
            timeout: 120000,                          // 120s — CogAgent inference takes ~31s
            maxContentLength: 50 * 1024 * 1024,       // 50 MB (screenshots are large)
            maxBodyLength:    50 * 1024 * 1024,
            headers: { 'Content-Type': 'application/json' },
        });
        const result = resp.data;
        result.action      = result.action      || 'fail';
        result.description = result.description || '';
        return result;
    } catch (error) {
        const msg = error.response
            ? `CogAgent HTTP ${error.response.status}: ${JSON.stringify(error.response.data).slice(0, 200)}`
            : `CogAgent unreachable: ${error.message}`;
        console.error('[cogagent-vision-act]', msg);
        return { action: 'fail', description: msg };
    }
});

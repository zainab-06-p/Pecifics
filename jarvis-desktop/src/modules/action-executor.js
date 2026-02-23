// ============================================
// JARVIS Action Executor Module
// Handles all desktop automation actions
// Pure PowerShell implementation - no native dependencies
// ============================================

const { exec, spawn } = require('child_process');
const fs = require('fs').promises;
const path = require('path');
const os = require('os');
const safetyGuard = require('./safety-guard');
const powerPointCOM = require('./powerpoint-com');
const wordCOM = require('./word-com');
const excelCOM = require('./excel-com');
const onenoteCOM = require('./onenote-com');
const publisherCOM = require('./publisher-com');
const fileManager = require('./file-manager');
const systemManager = require('./system-manager');
const osTasks = require('./os-tasks');
const browserAutomation = require('./browser-automation');

// PowerShell helper class for mouse/keyboard operations (loaded once)
const PS_HELPER_SCRIPT = `
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);
    
    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);
    
    public const int MOUSEEVENTF_LEFTDOWN = 0x02;
    public const int MOUSEEVENTF_LEFTUP = 0x04;
    public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
    public const int MOUSEEVENTF_RIGHTUP = 0x10;
    public const int MOUSEEVENTF_WHEEL = 0x0800;
}

[StructLayout(LayoutKind.Sequential)]
public struct POINT {
    public int X;
    public int Y;
}
"@
`;

let say = null;
let open = null;

// Try to load optional dependencies (lightweight ones only)
try {
    say = require('say');
} catch (e) {
    console.warn('say not available - using Windows SAPI for text-to-speech');
}

try {
    open = require('open');
} catch (e) {
    console.warn('open not available - URL opening may be limited');
}

// Application paths for Windows - comprehensive mapping with aliases
const APP_PATHS = {
    // Notepad
    'notepad': 'notepad.exe',
    'note pad': 'notepad.exe',
    'text editor': 'notepad.exe',
    
    // Calculator
    'calculator': 'calc.exe',
    'calc': 'calc.exe',
    'calci': 'calc.exe',
    
    // File Explorer
    'explorer': 'explorer.exe',
    'file explorer': 'explorer.exe',
    'files': 'explorer.exe',
    'my computer': 'explorer.exe',
    'this pc': 'explorer.exe',
    
    // Command Line
    'cmd': 'cmd.exe',
    'command prompt': 'cmd.exe',
    'terminal': 'cmd.exe',
    'powershell': 'powershell.exe',
    'ps': 'powershell.exe',
    'power shell': 'powershell.exe',
    
    // Browsers
    'chrome': 'chrome.exe',
    'google chrome': 'chrome.exe',
    'google': 'chrome.exe',
    'firefox': 'firefox.exe',
    'ff': 'firefox.exe',
    'mozilla': 'firefox.exe',
    'mozilla firefox': 'firefox.exe',
    'edge': 'msedge.exe',
    'microsoft edge': 'msedge.exe',
    'ms edge': 'msedge.exe',
    'brave': 'brave.exe',
    'opera': 'opera.exe',
    
    // Microsoft Office - Word
    'word': 'WINWORD.EXE',
    'microsoft word': 'WINWORD.EXE',
    'ms word': 'WINWORD.EXE',
    'msword': 'WINWORD.EXE',
    'winword': 'WINWORD.EXE',
    'doc': 'WINWORD.EXE',
    'document': 'WINWORD.EXE',
    'word processor': 'WINWORD.EXE',
    
    // Microsoft Office - Excel
    'excel': 'EXCEL.EXE',
    'microsoft excel': 'EXCEL.EXE',
    'ms excel': 'EXCEL.EXE',
    'msexcel': 'EXCEL.EXE',
    'spreadsheet': 'EXCEL.EXE',
    'xls': 'EXCEL.EXE',
    'xlsx': 'EXCEL.EXE',
    
    // Microsoft Office - PowerPoint
    'powerpoint': 'POWERPNT.EXE',
    'power point': 'POWERPNT.EXE',
    'microsoft powerpoint': 'POWERPNT.EXE',
    'ms powerpoint': 'POWERPNT.EXE',
    'mspowerpoint': 'POWERPNT.EXE',
    'ppt': 'POWERPNT.EXE',
    'pptx': 'POWERPNT.EXE',
    'slides': 'POWERPNT.EXE',
    'presentation': 'POWERPNT.EXE',
    
    // Microsoft Office - Outlook
    'outlook': 'OUTLOOK.EXE',
    'microsoft outlook': 'OUTLOOK.EXE',
    'ms outlook': 'OUTLOOK.EXE',
    'msoutlook': 'OUTLOOK.EXE',
    'mail': 'OUTLOOK.EXE',
    'email': 'OUTLOOK.EXE',
    
    // Microsoft Office - OneNote
    'onenote': 'ONENOTE.EXE',
    'one note': 'ONENOTE.EXE',
    'ms onenote': 'ONENOTE.EXE',
    
    // Microsoft Office - Access
    'access': 'MSACCESS.EXE',
    'ms access': 'MSACCESS.EXE',
    'microsoft access': 'MSACCESS.EXE',
    
    // VS Code
    'vscode': 'code',
    'vs code': 'code',
    'code': 'code',
    'visual studio code': 'code',
    
    // Visual Studio
    'visual studio': 'devenv.exe',
    'vs': 'devenv.exe',
    
    // System Tools
    'paint': 'mspaint.exe',
    'ms paint': 'mspaint.exe',
    'mspaint': 'mspaint.exe',
    'snipping tool': 'SnippingTool.exe',
    'snip': 'SnippingTool.exe',
    'screenshot': 'SnippingTool.exe',
    'task manager': 'taskmgr.exe',
    'taskmgr': 'taskmgr.exe',
    'task mgr': 'taskmgr.exe',
    'control panel': 'control.exe',
    'control': 'control.exe',
    'settings': 'ms-settings:',
    'windows settings': 'ms-settings:',
    
    // Media
    'vlc': 'vlc.exe',
    'media player': 'wmplayer.exe',
    'windows media player': 'wmplayer.exe',
    'wmp': 'wmplayer.exe',
    'spotify': 'spotify.exe',
    'itunes': 'iTunes.exe',
    
    // Communication
    'discord': 'discord.exe',
    'slack': 'slack.exe',
    'teams': 'teams.exe',
    'microsoft teams': 'teams.exe',
    'ms teams': 'teams.exe',
    'zoom': 'zoom.exe',
    'skype': 'skype.exe',
    'whatsapp': 'whatsapp.exe',
    'telegram': 'telegram.exe',
    
    // Development
    'git bash': 'git-bash.exe',
    'github': 'github.exe',
    'github desktop': 'github.exe',
    'postman': 'postman.exe',
    'sublime': 'sublime_text.exe',
    'sublime text': 'sublime_text.exe',
    'atom': 'atom.exe',
    'notepad++': 'notepad++.exe',
    'notepadplusplus': 'notepad++.exe',
    'npp': 'notepad++.exe',
    
    // Other common apps
    'steam': 'steam.exe',
    'obs': 'obs64.exe',
    'obs studio': 'obs64.exe',
    'photoshop': 'photoshop.exe',
    'adobe photoshop': 'photoshop.exe',
    'illustrator': 'illustrator.exe',
    'premiere': 'premiere.exe',
    'acrobat': 'acrobat.exe',
    'adobe acrobat': 'acrobat.exe',
    'pdf': 'acrobat.exe',
    'reader': 'AcroRd32.exe',
    'adobe reader': 'AcroRd32.exe',
    'pdf reader': 'AcroRd32.exe'
};

// Smart app name resolver - handles abbreviations, typos, and variations
function resolveAppName(input) {
    if (!input) return null;
    
    const normalized = input.toLowerCase().trim();
    
    // Direct match in APP_PATHS
    if (APP_PATHS[normalized]) {
        return APP_PATHS[normalized];
    }
    
    // Remove common prefixes/suffixes and try again
    const cleanedVariants = [
        normalized,
        normalized.replace(/^open\s+/, ''),       // "open word" -> "word"
        normalized.replace(/^launch\s+/, ''),     // "launch word" -> "word"
        normalized.replace(/^start\s+/, ''),      // "start word" -> "word"
        normalized.replace(/^run\s+/, ''),        // "run word" -> "word"
        normalized.replace(/\s+app$/, ''),        // "word app" -> "word"
        normalized.replace(/\s+application$/, ''), // "word application" -> "word"
        normalized.replace(/[.\-_]/g, ' '),       // "ms-word" -> "ms word"
        normalized.replace(/\s+/g, ''),           // "ms word" -> "msword"
    ];
    
    for (const variant of cleanedVariants) {
        if (APP_PATHS[variant]) {
            return APP_PATHS[variant];
        }
    }
    
    // Fuzzy matching - find closest match
    const appKeys = Object.keys(APP_PATHS);
    
    // Check if input contains any known app name
    for (const key of appKeys) {
        if (normalized.includes(key) || key.includes(normalized)) {
            return APP_PATHS[key];
        }
    }
    
    // Check for partial matches (at least 3 characters match)
    for (const key of appKeys) {
        const keyWords = key.split(/\s+/);
        const inputWords = normalized.split(/\s+/);
        
        for (const kw of keyWords) {
            for (const iw of inputWords) {
                if (kw.length >= 3 && iw.length >= 3) {
                    if (kw.startsWith(iw) || iw.startsWith(kw)) {
                        return APP_PATHS[key];
                    }
                }
            }
        }
    }
    
    // No match found, return original input (will try to run as-is)
    return input;
}

class ActionExecutor {
    constructor() {
        this.lastClickTime = 0;
        this.lastClickPos = { x: 0, y: 0 };
        // Track last opened app to re-focus before typing/pressing keys
        this.lastOpenedAppName = null;
        // Flag to stop execution
        this.shouldStop = false;
    }

    // Stop execution method
    stopExecution() {
        this.shouldStop = true;
        return { success: true, message: 'Stopping execution...' };
    }

    // Reset stop flag
    resetStopFlag() {
        this.shouldStop = false;
    }

    // Check if execution should stop
    isStopped() {
        return this.shouldStop;
    }

    // ============================================
    // File System Operations
    // ============================================

    async createFile(filePath, content = '') {
        try {
            // Ensure directory exists
            const dir = path.dirname(filePath);
            await fs.mkdir(dir, { recursive: true });
            
            // Write file
            await fs.writeFile(filePath, content, 'utf8');
            return { success: true, message: `Created file: ${filePath}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async createFolder(folderPath) {
        try {
            await fs.mkdir(folderPath, { recursive: true });
            return { success: true, message: `Created folder: ${folderPath}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async createPresentation(title, savePath, slidesContent) {
        try {
            const os = require('os');
            const path = require('path');
            // Resolve friendly paths
            let resolvedPath = savePath || 'Desktop';
            if (resolvedPath.toLowerCase() === 'desktop') resolvedPath = path.join(os.homedir(), 'Desktop');
            else if (resolvedPath.toLowerCase() === 'documents') resolvedPath = path.join(os.homedir(), 'Documents');
            else if (resolvedPath.toLowerCase() === 'downloads') resolvedPath = path.join(os.homedir(), 'Downloads');
            const presentationTitle = title || 'Presentation';
            const safeTitle = presentationTitle.replace(/[\\/:*?"<>|]/g, '_');
            const finalPath = path.join(resolvedPath, safeTitle.endsWith('.pptx') ? safeTitle : safeTitle + '.pptx');
            // Build slides array
            let slides = [];
            if (Array.isArray(slidesContent)) {
                slides = slidesContent;
            } else if (typeof slidesContent === 'string') {
                try { slides = JSON.parse(slidesContent); } catch { slides = [{ title: presentationTitle, content: slidesContent }]; }
            } else {
                slides = [{ title: presentationTitle, content: 'AI-generated presentation.' }];
            }
            if (slides.length === 0) slides = [{ title: presentationTitle, content: '' }];
            // Build PowerShell COM script
            let slideCmds = '';
            slides.forEach((s, i) => {
                const slideTitle = (s.title || 'Slide ' + (i + 1)).replace(/'/g, "''");
                const slideBody = (s.content || '').replace(/'/g, "''");
                if (i === 0) {
                    slideCmds += `\n$slide = $pres.Slides(1)\n$slide.Shapes(1).TextFrame.TextRange.Text = '${slideTitle}'\nif ($slide.Shapes.Count -gt 1) { $slide.Shapes(2).TextFrame.TextRange.Text = '${slideBody}' }\n`;
                } else {
                    slideCmds += `\n$slide = $pres.Slides.Add(${i + 1}, 1)\n$slide.Shapes(1).TextFrame.TextRange.Text = '${slideTitle}'\nif ($slide.Shapes.Count -gt 1) { $slide.Shapes(2).TextFrame.TextRange.Text = '${slideBody}' }\n`;
                }
            });
            const psScript = `$ppt = New-Object -ComObject PowerPoint.Application\n$ppt.Visible = 1\n$pres = $ppt.Presentations.Add()\n${slideCmds}\n$pres.SaveAs('${finalPath.replace(/\\/g, '\\\\')}')\n$ppt.Quit()`;
            const tmpFile = path.join(os.tmpdir(), 'create_pptx_' + Date.now() + '.ps1');
            require('fs').writeFileSync(tmpFile, psScript, 'utf8');
            const { execSync } = require('child_process');
            execSync(`powershell.exe -ExecutionPolicy Bypass -File "${tmpFile}"`, { timeout: 30000 });
            require('fs').unlinkSync(tmpFile);
            // Open the file to show the user the result
            const { shell } = require('electron');
            await shell.openPath(finalPath);
            return { success: true, message: `Presentation created and opened: ${finalPath}`, path: finalPath };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async deleteFile(filePath) {
        try {
            const stat = await fs.stat(filePath);
            if (stat.isDirectory()) {
                await fs.rmdir(filePath, { recursive: true });
            } else {
                await fs.unlink(filePath);
            }
            return { success: true, message: `Deleted: ${filePath}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async readFile(filePath) {
        try {
            const ext = path.extname(filePath).toLowerCase();
            // For Word docs, return paragraph list via COM
            if (ext === '.docx' || ext === '.doc') {
                return await wordCOM.readDocumentContent();
            }
            const content = await fs.readFile(filePath, 'utf8');
            return { success: true, content: content };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    // Find text inside any file and return matching lines with line numbers
    async findInFile(filePath, searchText) {
        try {
            const ext = path.extname(filePath).toLowerCase();
            if (ext === '.docx' || ext === '.doc') {
                // Word: search paragraphs via COM
                const result = await wordCOM.readDocumentContent();
                if (!result.success) return result;
                const matches = (result.paragraphs || []).filter(p =>
                    p.Text && p.Text.toLowerCase().includes(searchText.toLowerCase())
                );
                return { success: true, matches, count: matches.length,
                    message: `Found "${searchText}" in ${matches.length} paragraph(s)` };
            }
            if (ext === '.pptx' || ext === '.ppt') {
                return await powerPointCOM.findSlideByText(searchText);
            }
            const content = await fs.readFile(filePath, 'utf8');
            const lines = content.split('\n');
            const matches = [];
            lines.forEach((line, idx) => {
                if (line.toLowerCase().includes(searchText.toLowerCase())) {
                    matches.push({ line: idx + 1, text: line.trim() });
                }
            });
            return { success: true, matches, count: matches.length,
                message: `Found "${searchText}" in ${matches.length} line(s)` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    // Replace text inside any file (plain text, Word, PPT)
    async replaceInFile(filePath, searchText, replacementText, replaceAll = true) {
        try {
            const ext = path.extname(filePath).toLowerCase();
            if (ext === '.docx' || ext === '.doc') {
                return await wordCOM.findAndReplace(searchText, replacementText, replaceAll);
            }
            if (ext === '.pptx' || ext === '.ppt') {
                // For PPT: find the slide first, then update
                const found = await powerPointCOM.findSlideByText(searchText);
                if (!found.success || !found.slides || found.slides.length === 0) {
                    return { success: false, error: `"${searchText}" not found in any slide` };
                }
                const results = [];
                for (const s of found.slides) {
                    const r = await powerPointCOM.updateSlideText(
                        s.SlideNumber, searchText, replacementText
                    );
                    results.push(r);
                }
                return { success: true, message: `Updated ${results.length} slide(s)`, results };
            }
            // Plain text files (txt, md, csv, js, py, html, etc.)
            const content = await fs.readFile(filePath, 'utf8');
            const count = content.split(searchText).length - 1;
            if (count === 0) return { success: false, error: `"${searchText}" not found in file` };
            const updated = replaceAll
                ? content.split(searchText).join(replacementText)
                : content.replace(searchText, replacementText);
            await fs.writeFile(filePath, updated, 'utf8');
            return { success: true,
                message: `Replaced ${replaceAll ? count : 1} occurrence(s) of "${searchText}" with "${replacementText}"` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    // Append text to any file
    async appendToFile(filePath, content) {
        try {
            await fs.appendFile(filePath, '\n' + content, 'utf8');
            return { success: true, message: `Content appended to ${filePath}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async listDirectory(dirPath) {
        try {
            const items = await fs.readdir(dirPath, { withFileTypes: true });
            const result = items.map(item => ({
                name: item.name,
                type: item.isDirectory() ? 'folder' : 'file'
            }));
            return { success: true, items: result };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async copyFile(source, destination) {
        try {
            await fs.copyFile(source, destination);
            return { success: true, message: `Copied ${source} to ${destination}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async moveFile(source, destination) {
        try {
            await fs.rename(source, destination);
            return { success: true, message: `Moved ${source} to ${destination}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async renameFile(oldPath, newName) {
        try {
            const dir = path.dirname(oldPath);
            const newPath = path.join(dir, newName);
            await fs.rename(oldPath, newPath);
            return { success: true, message: `Renamed to ${newName}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    // ============================================
    // Application Control
    // ============================================

    async openApplication(appName) {
        return new Promise((resolve) => {
            // Use smart app name resolution
            const command = resolveAppName(appName);
            console.log(`Opening application: "${appName}" -> resolved to: "${command}"`);
            
            // Handle special cases (Windows Settings)
            if (command.startsWith('ms-settings:')) {
                exec(`start ${command}`, (error) => {
                    if (error) {
                        resolve({ success: false, error: error.message });
                    } else {
                        resolve({ success: true, message: `Opened ${appName}` });
                    }
                });
                return;
            }

            // Try multiple methods to open the application
            this.tryOpenApp(command, appName)
                .then(result => resolve(result))
                .catch(err => resolve({ success: false, error: err.message }));
        });
    }

    async tryOpenApp(command, originalName) {
        return new Promise(async (resolve) => {
            // Method 1: Direct start command
            exec(`start "" "${command}"`, async (error) => {
                if (!error) {
                    // Wait for app to start and get focus
                    await this.waitForAppAndFocus(command, originalName);

                    // Remember last opened app for focus retries
                    this.lastOpenedAppName = originalName;
                    resolve({ success: true, message: `Opened ${originalName}` });
                    return;
                }
                
                // Method 2: Try without quotes
                exec(`start ${command}`, async (err2) => {
                    if (!err2) {
                        await this.waitForAppAndFocus(command, originalName);

                        // Remember last opened app for focus retries
                        this.lastOpenedAppName = originalName;
                        resolve({ success: true, message: `Opened ${originalName}` });
                        return;
                    }
                    
                    // Method 3: Search in Start Menu
                    this.searchAndLaunchApp(originalName)
                        .then(async (result) => {
                            if (result.success) {
                                await this.waitForAppAndFocus(command, originalName);
                                this.lastOpenedAppName = originalName;
                            }
                            resolve(result);
                        })
                        .catch(() => {
                            resolve({ success: false, error: `Could not open ${originalName}` });
                        });
                });
            });
        });
    }

    async searchAndLaunchApp(appName) {
        return new Promise((resolve) => {
            // Use PowerShell to search Start Menu and launch
            const psScript = `
                $app = Get-StartApps | Where-Object { $_.Name -like '*${appName.replace(/'/g, "''")}*' } | Select-Object -First 1
                if ($app) {
                    Start-Process "explorer.exe" -ArgumentList "shell:AppsFolder\\$($app.AppID)"
                    Write-Output "Found and launched: $($app.Name)"
                } else {
                    Write-Error "App not found"
                    exit 1
                }
            `;
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: `Could not find ${appName} in Start Menu` });
                } else {
                    resolve({ success: true, message: stdout.trim() || `Opened ${appName}` });
                }
            });
        });
    }

    async waitForAppAndFocus(command, appName, maxWaitMs = 5000) {
        // Wait for the application window to appear and give it focus
        const startTime = Date.now();
        const processName = path.basename(command, '.exe').toUpperCase();
        let intervalCleared = false;
        let focusAttempts = 0;
        
        return new Promise((resolve) => {
            const checkInterval = setInterval(async () => {
                const elapsed = Date.now() - startTime;
                
                // STOP CONDITION: Prevent infinite loop
                if (elapsed >= maxWaitMs) {
                    if (!intervalCleared) {
                        clearInterval(checkInterval);
                        intervalCleared = true;
                    }
                    return;
                }
                
                // Try to activate the window
                try {
                    const focused = await this.focusApplicationWindow(appName, processName);
                    if (focused) {
                        focusAttempts++;
                        console.log(`✓ Focused ${appName} (attempt ${focusAttempts})`);
                    }
                } catch (e) {
                    // Ignore errors during focus attempts
                }
            }, 500);
            
            // GUARANTEED TIMEOUT: Force resolve after max wait
            setTimeout(() => {
                if (!intervalCleared) {
                    clearInterval(checkInterval);
                    intervalCleared = true;
                }
                console.log(`Focus complete for ${appName} after ${focusAttempts} successful attempts`);
                resolve();
            }, maxWaitMs + 100);
        });
    }

    async focusApplicationWindow(appName, processName) {
        return new Promise((resolve) => {
            // Use PowerShell to find and focus the window
            const psScript = `
                Add-Type @"
                using System;
                using System.Runtime.InteropServices;
                public class WindowHelper {
                    [DllImport("user32.dll")]
                    public static extern bool SetForegroundWindow(IntPtr hWnd);
                    [DllImport("user32.dll")]
                    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
                }
"@
                $processes = Get-Process | Where-Object { 
                    $_.MainWindowTitle -ne '' -and 
                    ($_.ProcessName -like '*${processName}*' -or $_.MainWindowTitle -like '*${appName}*')
                } | Select-Object -First 1
                
                if ($processes) {
                    [WindowHelper]::ShowWindow($processes.MainWindowHandle, 9)
                    [WindowHelper]::SetForegroundWindow($processes.MainWindowHandle)
                    Write-Output "Focused"
                }
            `;
            
            // Add timeout to prevent hanging
            const timeout = setTimeout(() => {
                resolve(false);
            }, 3000);
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`, (error) => {
                clearTimeout(timeout);
                resolve(error ? false : true);
            });
        });
    }

    async closeApplication(appName) {
        return new Promise((resolve) => {
            const appNameLower = appName.toLowerCase().trim();
            
            // Windows: use taskkill
            exec(`taskkill /IM "${appNameLower}.exe" /F`, (error) => {
                if (error) {
                    // Try without .exe
                    exec(`taskkill /IM "${appNameLower}" /F`, (err) => {
                        if (err) {
                            resolve({ success: false, error: `Could not close ${appName}` });
                        } else {
                            resolve({ success: true, message: `Closed ${appName}` });
                        }
                    });
                } else {
                    resolve({ success: true, message: `Closed ${appName}` });
                }
            });
        });
    }

    async openUrl(url) {
        try {
            if (open) {
                await open(url);
            } else {
                exec(`start "" "${url}"`);
            }
            return { success: true, message: `Opened ${url}` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    async downloadFile(url, filename) {
        return new Promise((resolve) => {
            const downloadPath = path.join(os.homedir(), 'Downloads', filename || 'presentation.pptx');
            console.log(`📥 Downloading ${filename} from ${url} to ${downloadPath}`);
            
            // Use PowerShell to download and open file
            const psScript = `
                try {
                    Invoke-WebRequest -Uri "${url}" -OutFile "${downloadPath}" -UseBasicParsing
                    Start-Process "${downloadPath}"
                    Write-Output "SUCCESS"
                } catch {
                    Write-Output "ERROR:$($_.Exception.Message)"
                }
            `.replace(/\n/g, ' ');
            
            exec(`powershell -Command "${psScript}"`, {timeout: 30000}, (error, stdout, stderr) => {
                const output = stdout.trim();
                if (output === 'SUCCESS') {
                    resolve({ success: true, message: `Downloaded and opened ${filename}`, path: downloadPath });
                } else if (error || stderr) {
                    resolve({ success: false, error: output || error?.message || stderr });
                } else {
                    resolve({ success: false, error: 'Download failed' });
                }
            });
        });
    }

    async searchWeb(query) {
        const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(query)}`;
        return this.openUrl(searchUrl);
    }

    // ============================================
    // Mouse Control (Pure PowerShell)
    // ============================================

    async moveMouse(x, y) {
        return new Promise((resolve) => {
            const psCommand = `${PS_HELPER_SCRIPT}; [Win32]::SetCursorPos(${x}, ${y})`;
            exec(`powershell -NoProfile -Command "${psCommand.replace(/"/g, '\"').replace(/\n/g, ' ')}"`, (error) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to move mouse' });
                } else {
                    resolve({ success: true, message: `Moved mouse to (${x}, ${y})` });
                }
            });
        });
    }

    async click(x, y, clickType = 'left', double = false) {
        return new Promise(async (resolve) => {
            try {
                // Move to position first
                if (x !== undefined && y !== undefined) {
                    await this.moveMouse(x, y);
                    await this.delay(100);
                }
                
                // Use a simpler PowerShell approach with proper error handling
                const clickScript = `
                    Add-Type @"
                        using System;
                        using System.Runtime.InteropServices;
                        public class MouseHelper {
                            [DllImport("user32.dll")]
                            public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
                            public const int LEFTDOWN = 0x02;
                            public const int LEFTUP = 0x04;
                            public const int RIGHTDOWN = 0x08;
                            public const int RIGHTUP = 0x10;
                        }
"@
                    if ('${clickType}' -eq 'right') {
                        [MouseHelper]::mouse_event([MouseHelper]::RIGHTDOWN, 0, 0, 0, 0)
                        Start-Sleep -Milliseconds 50
                        [MouseHelper]::mouse_event([MouseHelper]::RIGHTUP, 0, 0, 0, 0)
                    } else {
                        [MouseHelper]::mouse_event([MouseHelper]::LEFTDOWN, 0, 0, 0, 0)
                        Start-Sleep -Milliseconds 50
                        [MouseHelper]::mouse_event([MouseHelper]::LEFTUP, 0, 0, 0, 0)
                    }
                `;
                
                exec(`powershell -NoProfile -Command "${clickScript.replace(/\n/g, ' ').replace(/"/g, '\\"')}"`, { timeout: 5000 }, (error) => {
                    if (error) {
                        console.error('Click error:', error.message);
                        resolve({ success: false, error: 'Failed to click' });
                    } else {
                        resolve({ success: true, message: `Clicked at (${x}, ${y})` });
                    }
                });
            } catch (error) {
                console.error('Click exception:', error);
                resolve({ success: false, error: 'Failed to click' });
            }
        });
    }

    async scroll(direction, amount = 3) {
        return new Promise((resolve) => {
            // Scroll amount: positive = up, negative = down. Each "click" is 120 units
            const scrollValue = direction === 'up' ? (amount * 120) : -(amount * 120);
            const psCommand = `${PS_HELPER_SCRIPT}; [Win32]::mouse_event([Win32]::MOUSEEVENTF_WHEEL, 0, 0, ${scrollValue}, 0)`;
            
            exec(`powershell -NoProfile -Command "${psCommand.replace(/"/g, '\"').replace(/\n/g, ' ')}"`, (error) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to scroll' });
                } else {
                    resolve({ success: true, message: `Scrolled ${direction}` });
                }
            });
        });
    }

    async drag(startX, startY, endX, endY) {
        return new Promise(async (resolve) => {
            try {
                // Move to start position
                await this.moveMouse(startX, startY);
                await this.delay(100);
                
                // Mouse down
                const downCmd = `${PS_HELPER_SCRIPT}; [Win32]::mouse_event([Win32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)`;
                await this.execPowerShell(downCmd);
                await this.delay(100);
                
                // Move to end position (smooth drag)
                const steps = 10;
                for (let i = 1; i <= steps; i++) {
                    const currentX = Math.round(startX + (endX - startX) * (i / steps));
                    const currentY = Math.round(startY + (endY - startY) * (i / steps));
                    await this.moveMouse(currentX, currentY);
                    await this.delay(20);
                }
                
                // Mouse up
                const upCmd = `${PS_HELPER_SCRIPT}; [Win32]::mouse_event([Win32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)`;
                await this.execPowerShell(upCmd);
                
                resolve({ success: true, message: `Dragged from (${startX}, ${startY}) to (${endX}, ${endY})` });
            } catch (error) {
                resolve({ success: false, error: 'Drag failed' });
            }
        });
    }

    // Helper to execute PowerShell commands
    async execPowerShell(command) {
        return new Promise((resolve, reject) => {
            exec(`powershell -NoProfile -Command "${command.replace(/"/g, '\"').replace(/\n/g, ' ')}"`, (error, stdout) => {
                if (error) reject(error);
                else resolve(stdout);
            });
        });
    }

    // ============================================
    // Keyboard Control
    // ============================================

    async typeText(text, delay = 0) {
        if (!text) {
            return { success: false, error: 'No text provided' };
        }

        // First, ensure the target window has focus
        await this.delay(500);
        
        // Try to focus last opened app before typing
        await this.focusLastOpenedApp();
        await this.delay(300);

        // Use Windows PowerShell SendKeys with one retry after attempting focus
        const result = await this.typeTextPowerShell(text);
        if (result.success === false) {
            const refocused = await this.focusLastOpenedApp();
            if (refocused) {
                await this.delay(500);
                return await this.typeTextPowerShell(text);
            }
        }
        return result;
    }

    async typeTextPowerShell(text) {
        return new Promise((resolve) => {
            // Escape special SendKeys characters
            const escapedText = text
                .replace(/\+/g, '{+}')
                .replace(/\^/g, '{^}')
                .replace(/%/g, '{%}')
                .replace(/~/g, '{~}')
                .replace(/\(/g, '{(}')
                .replace(/\)/g, '{)}')
                .replace(/\[/g, '{[}')
                .replace(/\]/g, '{]}')
                .replace(/\{/g, '{{}')
                .replace(/\}/g, '{}}')
                .replace(/"/g, '""')
                .replace(/\r?\n/g, '{ENTER}')
                .replace(/\t/g, '{TAB}');
            
            // Use SendWait with explicit focus check
            const psCommand = `
                Add-Type -AssemblyName System.Windows.Forms;
                Add-Type -AssemblyName System.Runtime.InteropServices;
                $sig = '[DllImport(\\"user32.dll\\")] public static extern IntPtr GetForegroundWindow();';
                $type = Add-Type -MemberDefinition $sig -Name WinAPI -Namespace Win32 -PassThru;
                $hwnd = $type::GetForegroundWindow();
                if ($hwnd -ne [IntPtr]::Zero) {
                    [System.Windows.Forms.SendKeys]::SendWait(\\"${escapedText}\\");
                    Write-Output 'OK';
                } else {
                    throw 'No window focused';
                }
            `.replace(/\n/g, ' ');
            
            exec(`powershell -NoProfile -Command "${psCommand}"`, { timeout: 30000 }, (error, stdout, stderr) => {
                if (error || stderr) {
                    console.error('Typing error:', error?.message || stderr);
                    resolve({ success: false, error: 'Typing failed. Make sure a text field is focused.' });
                } else if (stdout && stdout.includes('OK')) {
                    resolve({ success: true, message: `Typed: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"` });
                } else {
                    resolve({ success: false, error: 'Typing failed. No window focused.' });
                }
            });
        });
    }

    // Type text directly into an application (with focus management)
    async typeIntoApp(appName, text, delay = 0) {
        // First focus the app
        const resolved = resolveAppName(appName);
        const processName = path.basename(resolved, '.exe').toUpperCase();
        
        await this.focusApplicationWindow(appName, processName);
        await this.delay(500); // Wait for focus
        
        // Then type
        return this.typeText(text, delay);
    }

    // Helper delay function
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async pressKey(key) {
        // Ensure window is focused before pressing key
        await this.delay(200);
        await this.focusLastOpenedApp();
        await this.delay(200);
        
        // Use Windows PowerShell SendKeys for all key presses
        return new Promise(async (resolve) => {
            const keyLower = key.toLowerCase();
            
            // Comprehensive key mapping for SendKeys format
            const keyMap = {
                // Special keys
                'enter': '{ENTER}',
                'return': '{ENTER}',
                'tab': '{TAB}',
                'backspace': '{BACKSPACE}',
                'back': '{BACKSPACE}',
                'delete': '{DELETE}',
                'del': '{DELETE}',
                'escape': '{ESC}',
                'esc': '{ESC}',
                'space': ' ',
                'spacebar': ' ',
                
                // Arrow keys
                'up': '{UP}',
                'down': '{DOWN}',
                'left': '{LEFT}',
                'right': '{RIGHT}',
                
                // Navigation
                'home': '{HOME}',
                'end': '{END}',
                'pageup': '{PGUP}',
                'pagedown': '{PGDN}',
                'pgup': '{PGUP}',
                'pgdn': '{PGDN}',
                'insert': '{INSERT}',
                'ins': '{INSERT}',
                
                // Function keys
                'f1': '{F1}', 'f2': '{F2}', 'f3': '{F3}', 'f4': '{F4}',
                'f5': '{F5}', 'f6': '{F6}', 'f7': '{F7}', 'f8': '{F8}',
                'f9': '{F9}', 'f10': '{F10}', 'f11': '{F11}', 'f12': '{F12}',
                
                // Common shortcuts
                'ctrl+s': '^s',
                'ctrl+c': '^c',
                'ctrl+v': '^v',
                'ctrl+x': '^x',
                'ctrl+z': '^z',
                'ctrl+y': '^y',
                'ctrl+a': '^a',
                'ctrl+f': '^f',
                'ctrl+n': '^n',
                'ctrl+o': '^o',
                'ctrl+p': '^p',
                'ctrl+w': '^w',
                'ctrl+shift+s': '^+s',
                'ctrl+shift+n': '^+n',
                'alt+f4': '%{F4}',
                'alt+tab': '%{TAB}',
                'win+d': '^{ESC}d',
                'win+e': '^{ESC}e',
                'win+r': '^{ESC}r'
            };
            
            let sendKey;
            
            // Check if it's a known mapping
            if (keyMap[keyLower]) {
                sendKey = keyMap[keyLower];
            } else if (keyLower.includes('+')) {
                // Parse custom key combination
                sendKey = this.parseKeyCombo(keyLower);
            } else {
                // Single character or unknown key
                sendKey = key.length === 1 ? key : `{${key.toUpperCase()}}`;
            }
            
            // Use improved command with focus check
            const psCommand = `
                Add-Type -AssemblyName System.Windows.Forms;
                Add-Type -AssemblyName System.Runtime.InteropServices;
                $sig = '[DllImport(\\"user32.dll\\")] public static extern IntPtr GetForegroundWindow();';
                $type = Add-Type -MemberDefinition $sig -Name WinAPI2 -Namespace Win32Key -PassThru -ErrorAction SilentlyContinue;
                if (!$type) { $type = [Win32Key.WinAPI2] }
                $hwnd = $type::GetForegroundWindow();
                if ($hwnd -ne [IntPtr]::Zero) {
                    [System.Windows.Forms.SendKeys]::SendWait(\\"${sendKey}\\");
                    Write-Output 'OK';
                } else {
                    throw 'No window focused';
                }
            `.replace(/\n/g, ' ');
            
            exec(`powershell -NoProfile -Command "${psCommand}"`, { timeout: 5000 }, async (error, stdout, stderr) => {
                if (error || stderr || !stdout?.includes('OK')) {
                    console.error('Key press error:', error?.message || stderr);
                    // Retry after attempting focus of last opened app
                    const refocused = await this.focusLastOpenedApp();
                    if (refocused) {
                        await this.delay(300);
                        exec(`powershell -NoProfile -Command "${psCommand}"`, { timeout: 5000 }, (err2, stdout2) => {
                            if (err2 || !stdout2?.includes('OK')) {
                                resolve({ success: false, error: 'Key press failed' });
                            } else {
                                resolve({ success: true, message: `Pressed ${key}` });
                            }
                        });
                        return;
                    }
                    resolve({ success: false, error: 'Key press failed' });
                } else {
                    resolve({ success: true, message: `Pressed ${key}` });
                }
            });
        });
    }

    // Try to bring last opened app to foreground for retries
    async focusLastOpenedApp() {
        if (!this.lastOpenedAppName) return false;
        try {
            const resolved = resolveAppName(this.lastOpenedAppName);
            const processName = path.basename(resolved, '.exe').toUpperCase();
            const focused = await this.focusApplicationWindow(this.lastOpenedAppName, processName);
            return focused;
        } catch (e) {
            return false;
        }
    }

    // Parse key combinations like "ctrl+shift+s" into SendKeys format
    parseKeyCombo(combo) {
        const parts = combo.toLowerCase().split('+');
        let result = '';
        let mainKey = '';
        
        for (const part of parts) {
            const p = part.trim();
            if (p === 'ctrl' || p === 'control') {
                result += '^';
            } else if (p === 'alt') {
                result += '%';
            } else if (p === 'shift') {
                result += '+';
            } else if (p === 'win' || p === 'windows' || p === 'meta') {
                result += '^{ESC}'; // Windows key approximation
            } else {
                mainKey = p;
            }
        }
        
        // Add the main key
        if (mainKey.length === 1) {
            result += mainKey;
        } else {
            result += `{${mainKey.toUpperCase()}}`;
        }
        
        return result;
    }

    async holdKey(key, action = 'down') {
        // SendKeys doesn't support hold, so we simulate with key press
        // For most use cases, just pressing the key is sufficient
        return this.pressKey(key);
    }

    // ============================================
    // System Operations
    // ============================================

    async runCommand(command) {
        return new Promise((resolve) => {
            exec(command, { shell: true }, (error, stdout, stderr) => {
                if (error) {
                    resolve({ 
                        success: false, 
                        error: error.message,
                        stderr: stderr 
                    });
                } else {
                    resolve({ 
                        success: true, 
                        output: stdout,
                        stderr: stderr 
                    });
                }
            });
        });
    }

    speak(message, voice = null, speed = 1.0) {
        return new Promise((resolve) => {
            if (!say) {
                // Fallback to Windows SAPI
                const escapedMessage = message.replace(/"/g, '\\"');
                exec(`powershell -Command "Add-Type -AssemblyName System.Speech; $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer; $synth.Speak('${escapedMessage}')"`, 
                    (error) => {
                        if (error) {
                            resolve({ success: false, error: 'Text-to-speech failed' });
                        } else {
                            resolve({ success: true, message: 'Spoke message' });
                        }
                    }
                );
                return;
            }
            
            say.speak(message, voice, speed, (error) => {
                if (error) {
                    resolve({ success: false, error: error.message });
                } else {
                    resolve({ success: true, message: 'Spoke message' });
                }
            });
        });
    }

    // ============================================
    // Composite Actions (Multi-step automation)
    // ============================================

    // Open app and type text into it
    async openAppAndType(appName, text, waitTime = 3000) {
        const openResult = await this.openApplication(appName);
        if (!openResult.success) return openResult;
        
        await this.delay(waitTime); // Wait for app to fully load
        
        const typeResult = await this.typeText(text);
        return {
            success: typeResult.success,
            message: `Opened ${appName} and typed text`
        };
    }

    // PowerPoint: Create new slide
    async pptNewSlide() {
        await this.pressKey('ctrl+m');
        await this.delay(500);
        return { success: true, message: 'Created new slide' };
    }

    // PowerPoint: Add title to current slide
    async pptAddTitle(title) {
        // Click on title placeholder (usually top center)
        await this.pressKey('ctrl+shift+enter'); // Select title placeholder
        await this.delay(300);
        await this.typeText(title);
        return { success: true, message: `Added title: ${title}` };
    }

    // PowerPoint: Add content/body text
    async pptAddContent(content) {
        await this.pressKey('tab'); // Move to content area
        await this.delay(200);
        await this.typeText(content);
        return { success: true, message: 'Added content' };
    }

    // Word: Type paragraph
    async wordTypeParagraph(text) {
        await this.typeText(text);
        await this.pressKey('enter');
        await this.pressKey('enter');
        return { success: true, message: 'Typed paragraph' };
    }

    // Word: Insert heading
    async wordInsertHeading(text, level = 1) {
        // Apply heading style
        await this.pressKey(`ctrl+alt+${level}`);
        await this.delay(200);
        await this.typeText(text);
        await this.pressKey('enter');
        return { success: true, message: `Inserted heading level ${level}` };
    }

    // Save current document
    async saveDocument(filename = null) {
        if (filename) {
            // Save As
            await this.pressKey('ctrl+shift+s');
            await this.delay(500);
            await this.typeText(filename);
            await this.delay(200);
            await this.pressKey('enter');
        } else {
            // Quick save
            await this.pressKey('ctrl+s');
        }
        await this.delay(500);
        return { success: true, message: filename ? `Saved as ${filename}` : 'Saved document' };
    }

    // Select all and copy
    async selectAllCopy() {
        await this.pressKey('ctrl+a');
        await this.delay(200);
        await this.pressKey('ctrl+c');
        return { success: true, message: 'Selected all and copied' };
    }

    // Paste
    async paste() {
        await this.pressKey('ctrl+v');
        return { success: true, message: 'Pasted from clipboard' };
    }

    // Undo
    async undo() {
        await this.pressKey('ctrl+z');
        return { success: true, message: 'Undone last action' };
    }

    // Redo
    async redo() {
        await this.pressKey('ctrl+y');
        return { success: true, message: 'Redone action' };
    }

    async getMousePosition() {
        return new Promise((resolve) => {
            const psScript = `${PS_HELPER_SCRIPT}; $point = New-Object POINT; [Win32]::GetCursorPos([ref]$point) | Out-Null; Write-Output "$($point.X),$($point.Y)"`;
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/"/g, '\\"').replace(/\\n/g, ' ')}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to get mouse position' });
                } else {
                    const [x, y] = stdout.trim().split(',').map(Number);
                    resolve({ success: true, x, y });
                }
            });
        });
    }

    async getScreenSize() {
        return new Promise((resolve) => {
            const psScript = `Add-Type -AssemblyName System.Windows.Forms; $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; Write-Output "$($screen.Width),$($screen.Height)"`;
            
            exec(`powershell -NoProfile -Command "${psScript}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to get screen size' });
                } else {
                    const [width, height] = stdout.trim().split(',').map(Number);
                    resolve({ success: true, width, height });
                }
            });
        });
    }
    
    // ============================================
    // Interactive Choice Detection
    // ============================================
    
    async detectInstalledBrowsers() {
        return new Promise((resolve) => {
            const psScript = `
                $browsers = @()
                
                # Check Chrome
                if (Test-Path "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe") {
                    $browsers += "Chrome"
                }
                if (Test-Path "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe") {
                    $browsers += "Chrome"
                }
                
                # Check Edge
                if (Test-Path "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe") {
                    $browsers += "Edge"
                }
                if (Test-Path "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe") {
                    $browsers += "Edge"
                }
                
                # Check Firefox
                if (Test-Path "C:\\Program Files\\Mozilla Firefox\\firefox.exe") {
                    $browsers += "Firefox"
                }
                if (Test-Path "C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe") {
                    $browsers += "Firefox"
                }
                
                # Check Brave
                if (Test-Path "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe") {
                    $browsers += "Brave"
                }
                
                $browsers | Select-Object -Unique | ForEach-Object { Write-Output $_ }
            `;
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/\n/g, ' ')}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to detect browsers' });
                } else {
                    const browsers = stdout.trim().split('\n').filter(b => b.trim());
                    resolve({ 
                        success: true, 
                        browsers: browsers.length > 0 ? browsers : ['Edge'] // Default to Edge if none found
                    });
                }
            });
        });
    }
    
    async detectChromeProfiles() {
        return new Promise((resolve) => {
            const psScript = `
                $chromeUserData = "$env:LOCALAPPDATA\\Google\\Chrome\\User Data"
                $profiles = @()
                
                if (Test-Path $chromeUserData) {
                    Get-ChildItem -Path $chromeUserData -Directory | Where-Object {
                        $_.Name -match '^(Default|Profile \\d+)$'
                    } | ForEach-Object {
                        $prefsPath = Join-Path $_.FullName "Preferences"
                        if (Test-Path $prefsPath) {
                            $prefs = Get-Content $prefsPath -Raw | ConvertFrom-Json
                            $name = if ($prefs.profile.name) { $prefs.profile.name } else { $_.Name }
                            $profiles += "$name"
                        } else {
                            $profiles += $_.Name
                        }
                    }
                }
                
                if ($profiles.Count -eq 0) {
                    $profiles += "Default"
                }
                
                $profiles | ForEach-Object { Write-Output $_ }
            `;
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/\n/g, ' ')}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to detect Chrome profiles' });
                } else {
                    const profiles = stdout.trim().split('\n').filter(p => p.trim());
                    resolve({ 
                        success: true, 
                        profiles: profiles.length > 0 ? profiles : ['Default']
                    });
                }
            });
        });
    }
    
    async detectEdgeProfiles() {
        return new Promise((resolve) => {
            const psScript = `
                $edgeUserData = "$env:LOCALAPPDATA\\Microsoft\\Edge\\User Data"
                $profiles = @()
                
                if (Test-Path $edgeUserData) {
                    Get-ChildItem -Path $edgeUserData -Directory | Where-Object {
                        $_.Name -match '^(Default|Profile \\d+)$'
                    } | ForEach-Object {
                        $prefsPath = Join-Path $_.FullName "Preferences"
                        if (Test-Path $prefsPath) {
                            $prefs = Get-Content $prefsPath -Raw | ConvertFrom-Json
                            $name = if ($prefs.profile.name) { $prefs.profile.name } else { $_.Name }
                            $profiles += "$name"
                        } else {
                            $profiles += $_.Name
                        }
                    }
                }
                
                if ($profiles.Count -eq 0) {
                    $profiles += "Default"
                }
                
                $profiles | ForEach-Object { Write-Output $_ }
            `;
            
            exec(`powershell -NoProfile -Command "${psScript.replace(/\n/g, ' ')}"`, (error, stdout) => {
                if (error) {
                    resolve({ success: false, error: 'Failed to detect Edge profiles' });
                } else {
                    const profiles = stdout.trim().split('\n').filter(p => p.trim());
                    resolve({ 
                        success: true, 
                        profiles: profiles.length > 0 ? profiles : ['Default']
                    });
                }
            });
        });
    }

    // ============================================
    // Execute Action by Name (WITH SAFETY CHECK)
    // ============================================

    async execute(actionName, params = {}) {
        // 🛡️ SAFETY CHECK FIRST
        const safetyCheck = safetyGuard.validateAction(actionName, params);
        
        if (!safetyCheck.safe) {
            console.warn('Action blocked by safety guard:', safetyCheck.reason);
            return { 
                success: false, 
                error: safetyCheck.reason,
                blocked: true 
            };
        }

        // Log warnings if any
        if (safetyCheck.warnings.length > 0) {
            console.warn('Safety warnings:', safetyCheck.warnings);
        }

        const actions = {
            // File operations
            'create_file': () => this.createFile(params.file_path, params.content || ''),
            'create-file': () => this.createFile(params.file_path, params.content || ''),
            'create_folder': () => this.createFolder(params.folder_path),
            'create-folder': () => this.createFolder(params.folder_path),
            'delete_file': () => this.deleteFile(params.file_path),
            'delete-file': () => this.deleteFile(params.file_path),
            'read_file': () => this.readFile(params.file_path),
            'read-file': () => this.readFile(params.file_path),
            'list_directory': () => this.listDirectory(params.path),
            'list-directory': () => this.listDirectory(params.path),
            'copy_file': () => this.copyFile(params.source, params.destination),
            'copy-file': () => this.copyFile(params.source, params.destination),
            'move_file': () => this.moveFile(params.source, params.destination),
            'move-file': () => this.moveFile(params.source, params.destination),
            'rename_file': () => this.renameFile(params.old_path, params.new_name),
            'rename-file': () => this.renameFile(params.old_path, params.new_name),

            // Find / modify existing file content (any file type)
            'find_in_file': () => this.findInFile(params.file_path, params.search_text),
            'find-in-file': () => this.findInFile(params.file_path, params.search_text),
            'replace_in_file': () => this.replaceInFile(params.file_path, params.search_text, params.replacement_text, params.replace_all !== false),
            'replace-in-file': () => this.replaceInFile(params.file_path, params.search_text, params.replacement_text, params.replace_all !== false),
            'append_to_file': () => this.appendToFile(params.file_path, params.content),
            'append-to-file': () => this.appendToFile(params.file_path, params.content),
            'search_files': () => fileManager.searchFiles(params.pattern, params.location),
            'search-files': () => fileManager.searchFiles(params.pattern, params.location),
            'search_file_content': () => fileManager.searchFileContent(params.search_text, params.location, params.file_pattern),
            
            // Application control
            'open_application': () => this.openApplication(params.app_name),
            'open-app': () => this.openApplication(params.app_name),
            'open_app': () => this.openApplication(params.app_name),
            'launch_app': () => this.openApplication(params.app_name),
            'launch_application': () => this.openApplication(params.app_name),
            'close_application': () => this.closeApplication(params.app_name),
            'close-app': () => this.closeApplication(params.app_name),
            'close_app': () => this.closeApplication(params.app_name),
            'open_url': () => this.openUrl(params.url),
            'open-url': () => this.openUrl(params.url),
            'download_file': () => this.downloadFile(params.url, params.filename),
            'download-file': () => this.downloadFile(params.url, params.filename),
            'search_web': () => this.searchWeb(params.query),
            'search-web': () => this.searchWeb(params.query),
            'focus_app': () => this.focusApp(params.app_name),
            'focus_application': () => this.focusApp(params.app_name),
            
            // Mouse control
            'move_mouse': () => this.moveMouse(params.x, params.y),
            'move-mouse': () => this.moveMouse(params.x, params.y),
            'click_at': () => this.click(params.x, params.y, params.click_type, params.double),
            'click': () => this.click(params.x, params.y, params.click_type, params.double),
            'scroll': () => this.scroll(params.direction, params.amount),
            'drag': () => this.drag(params.start_x, params.start_y, params.end_x, params.end_y),
            
            // Keyboard control
            'type_text': () => this.typeText(params.text, params.delay),
            'type-text': () => this.typeText(params.text, params.delay),
            'type': () => this.typeText(params.text, params.delay),
            'type_into_app': () => this.typeIntoApp(params.app_name, params.text, params.delay),
            'type-into-app': () => this.typeIntoApp(params.app_name, params.text, params.delay),
            'press_key': () => this.pressKey(params.key),
            'press-key': () => this.pressKey(params.key),
            'press': () => this.pressKey(params.key),
            'hotkey': () => this.pressKey(params.key),
            'hold_key': () => this.holdKey(params.key, params.action),
            
            // Timing control
            'wait': () => this.wait(params.duration || params.ms || params.seconds * 1000 || 1000),
            'delay': () => this.wait(params.duration || params.ms || params.seconds * 1000 || 1000),
            'sleep': () => this.wait(params.duration || params.ms || params.seconds * 1000 || 1000),
            
            // System
            'run_command': () => this.runCommand(params.command),
            'run-command': () => this.runCommand(params.command),
            'speak': () => this.speak(params.message, params.voice, params.speed),
            
            // Composite actions (multi-step)
            'open_app_and_type': () => this.openAppAndType(params.app_name, params.text, params.wait_time),
            'open_and_type': () => this.openAppAndType(params.app_name, params.text, params.wait_time),
            
            // PowerPoint specific
            'ppt_new_slide': () => this.pptNewSlide(),
            'ppt_add_title': () => this.pptAddTitle(params.title || params.text),
            'ppt_add_content': () => this.pptAddContent(params.content || params.text),
            'powerpoint_new_slide': () => this.pptNewSlide(),
            'powerpoint_add_title': () => this.pptAddTitle(params.title || params.text),
            'powerpoint_add_content': () => this.pptAddContent(params.content || params.text),
            'create_presentation': () => this.createPresentation(params.title, params.save_path, params.slides_content),

            // PowerPoint COM automation (themes, animations, layouts)
            'ppt_apply_theme': () => this.ppt_apply_theme(params.theme_name),
            'ppt_add_animation': () => this.ppt_add_animation(params.animation_type, params.apply_to_all),
            'ppt_change_layout': () => this.ppt_change_layout(params.layout_name),

            // PowerPoint COM — find & modify existing slides
            'ppt_find_slide': () => powerPointCOM.findSlideByText(params.search_text),
            'ppt_get_slide_content': () => powerPointCOM.getSlideContent(params.slide_number),
            'ppt_update_slide_text': () => powerPointCOM.updateSlideText(params.slide_number, params.old_text, params.new_text),
            
            // Word COM automation - pass all params as options object
            'word_create_document': () => wordCOM.createDocument(params.title, params.content),
            'word_open_document': () => wordCOM.openDocument(params.filepath),
            'word_read_content': () => wordCOM.readDocumentContent(),
            'word_find_replace': () => wordCOM.findAndReplace(params.search_text, params.replacement_text, params.replace_all !== false),
            'word_add_paragraph': () => wordCOM.addParagraph(params.text, params),
            'word_add_heading': () => wordCOM.addHeading(params.text, params.level || 1, params),
            'word_insert_table': () => wordCOM.insertTable(params.rows, params.cols),
            'word_apply_theme': () => wordCOM.applyTheme(params.theme_name),
            'word_save': () => wordCOM.saveDocument(params.filename),
            'word_remove_table_borders': () => wordCOM.removeTableBorders(params.table_number),
            'word_change_font': () => wordCOM.changeFont(params.font_name, params.font_size, params.start_paragraph, params.end_paragraph),
            'word_clear_formatting': () => wordCOM.clearFormatting(params.start_paragraph, params.end_paragraph),
            'word_change_color': () => wordCOM.changeColor(params.color, params.start_paragraph, params.end_paragraph),
            
            // Excel COM automation
            'excel_create_workbook': () => excelCOM.createWorkbook(),
            'excel_open_workbook': () => excelCOM.openWorkbook(params.filepath),
            'excel_write_cell': () => excelCOM.writeCell(params.row, params.col, params.value),
            'excel_write_data': () => excelCOM.writeData(params.data),
            'excel_add_worksheet': () => excelCOM.addWorksheet(params.name),
            'excel_create_chart': () => excelCOM.createChart(params.chart_type),
            'excel_format_cell': () => excelCOM.formatCell(params.row, params.col, params),
            'excel_format_range': () => excelCOM.formatRange(params.start_row, params.start_col, params.end_row, params.end_col, params),
            'excel_autofit_columns': () => excelCOM.autoFitColumns(),
            'excel_save': () => excelCOM.saveWorkbook(params.filename),
            'excel_remove_borders': () => excelCOM.removeBorders(params.start_row, params.start_col, params.end_row, params.end_col),
            'excel_clear_formatting': () => excelCOM.clearFormatting(params.start_row, params.start_col, params.end_row, params.end_col),
            'excel_remove_background': () => excelCOM.removeBackgroundColor(params.start_row, params.start_col, params.end_row, params.end_col),
            'excel_change_font': () => excelCOM.changeFont(params.font_name, params.font_size, params.start_row, params.start_col, params.end_row, params.end_col),
            
            // OneNote COM automation
            'onenote_open': () => onenoteCOM.openOneNote(),
            'onenote_create_page': () => onenoteCOM.createPage(params.title, params.section),
            'onenote_add_content': () => onenoteCOM.addContent(params.content),
            
            // Publisher COM automation
            'publisher_create': () => publisherCOM.createPublication(params.template_type),
            'publisher_add_textbox': () => publisherCOM.addTextBox(params.text, params.left, params.top, params.width, params.height),
            'publisher_add_page': () => publisherCOM.addPage(),
            'publisher_save': () => publisherCOM.savePublication(params.filename),
            
            // Paint automation (cursor-based)
            'paint_open': () => this.openApplication('mspaint.exe'),
            'paint_draw_line': () => this.paintDrawLine(params.start_x, params.start_y, params.end_x, params.end_y),
            'paint_select_tool': () => this.paintSelectTool(params.tool),
            'paint_select_color': () => this.paintSelectColor(params.color),
            
            // Whiteboard automation (cursor-based)
            'whiteboard_open': () => this.openApplication('Microsoft Whiteboard'),
            
            // File system utilities
            'check_file_exists': () => this.checkFileExists(params.filepath),
            'search_files': () => this.searchFiles(params.search_term, params.file_type, params.search_location),
            'generate_unique_filename': () => this.generateUniqueFilename(params.filepath, params.content_description),
            
            // Word specific (backward compatibility)
            'word_type_paragraph': () => this.wordTypeParagraph(params.text),
            'word_insert_heading': () => this.wordInsertHeading(params.text, params.level),
            
            // Common document actions
            'save': () => this.saveDocument(params.filename),
            'save_document': () => this.saveDocument(params.filename),
            'save_as': () => this.saveDocument(params.filename),
            'select_all': () => this.pressKey('ctrl+a'),
            'copy': () => this.pressKey('ctrl+c'),
            'paste': () => this.paste(),
            'cut': () => this.pressKey('ctrl+x'),
            'undo': () => this.undo(),
            'redo': () => this.redo(),
            'select_all_copy': () => this.selectAllCopy(),
            
            // Advanced File Management
            'search_files_advanced': () => fileManager.searchFiles(params.pattern, params.location, params.max_results),
            'search_file_content': () => fileManager.searchFileContent(params.search_text, params.location, params.file_pattern),
            'delete_files': () => fileManager.deleteFiles(params.file_paths),
            'delete_multiple_files': () => fileManager.deleteFiles(params.file_paths),
            'copy_files': () => fileManager.copyFiles(params.file_paths, params.destination),
            'copy_multiple_files': () => fileManager.copyFiles(params.file_paths, params.destination),
            'move_files': () => fileManager.moveFiles(params.file_paths, params.destination),
            'move_multiple_files': () => fileManager.moveFiles(params.file_paths, params.destination),
            'create_copies': () => fileManager.createCopies(params.source_file, params.new_names, params.destination),
            'get_files_by_type': () => fileManager.getFilesByType(params.file_type, params.location, params.max_results),
            'get_file_info': () => fileManager.getFileInfo(params.file_path),
            'open_file': () => fileManager.openFile(params.file_path),
            'show_in_explorer': () => fileManager.showInExplorer(params.file_path),
            
            // Application Management
            'search_applications': () => fileManager.searchApplications(params.app_name),
            'search_apps': () => fileManager.searchApplications(params.app_name),
            'find_application': () => fileManager.searchApplications(params.app_name),
            'get_running_apps': () => fileManager.getRunningApplications(),
            'get_running_applications': () => fileManager.getRunningApplications(),
            'list_running_apps': () => fileManager.getRunningApplications(),
            'close_app_by_id': () => fileManager.closeApplication(params.identifier),
            'kill_process': () => fileManager.closeApplication(params.identifier),
            'uninstall_application': () => fileManager.uninstallApplication(params.app_name, params.force || false),
            'delete_application': () => fileManager.uninstallApplication(params.app_name, params.force || false),
            'remove_application': () => fileManager.uninstallApplication(params.app_name, params.force || false),
            'move_application': () => fileManager.moveApplication(params.app_name, params.new_location, params.update_registry !== false),
            'relocate_application': () => fileManager.moveApplication(params.app_name, params.new_location, params.update_registry !== false),
            
            // System Management
            'clear_temp_files': () => systemManager.clearTempFiles(params.include_cache !== false),
            'cleanup_temp': () => systemManager.clearTempFiles(params.include_cache !== false),
            'delete_temp_files': () => systemManager.clearTempFiles(params.include_cache !== false),
            'set_wallpaper': () => systemManager.setWallpaper(params.image_path),
            'change_wallpaper': () => systemManager.setWallpaper(params.image_path),
            'update_background': () => systemManager.setWallpaper(params.image_path),
            'toggle_wifi': () => systemManager.toggleWiFi(params.enable),
            'enable_wifi': () => systemManager.toggleWiFi(true),
            'disable_wifi': () => systemManager.toggleWiFi(false),
            'wifi_on': () => systemManager.toggleWiFi(true),
            'wifi_off': () => systemManager.toggleWiFi(false),
            'toggle_bluetooth': () => systemManager.toggleBluetooth(params.enable),
            'enable_bluetooth': () => systemManager.toggleBluetooth(true),
            'disable_bluetooth': () => systemManager.toggleBluetooth(false),
            'bluetooth_on': () => systemManager.toggleBluetooth(true),
            'bluetooth_off': () => systemManager.toggleBluetooth(false),
            'set_brightness': () => systemManager.setBrightness(params.brightness),
            'change_brightness': () => systemManager.setBrightness(params.brightness),
            'adjust_brightness': () => systemManager.setBrightness(params.brightness),
            'get_battery_status': () => systemManager.getBatteryStatus(),
            'check_battery': () => systemManager.getBatteryStatus(),
            'battery_info': () => systemManager.getBatteryStatus(),
            'set_resolution': () => systemManager.setResolution(params.width, params.height),
            'change_resolution': () => systemManager.setResolution(params.width, params.height),
            'screen_resolution': () => systemManager.setResolution(params.width, params.height),
            'get_system_info': () => systemManager.getSystemInfo(),
            'system_info': () => systemManager.getSystemInfo(),
            'computer_info': () => systemManager.getSystemInfo(),
            // Screenshot / vision — these are intercepted in renderer.js; here just acknowledge
            'get_screenshot': () => Promise.resolve({ success: true, message: 'Screenshot captured for analysis' }),
            'describe_screen': () => Promise.resolve({ success: true, message: 'Screen analysis requested' }),
            'analyze_screen': () => Promise.resolve({ success: true, message: 'Screen analysis requested' }),
            'what_is_on_screen': () => Promise.resolve({ success: true, message: 'Screen analysis requested' }),
            'toggle_night_light': () => systemManager.toggleNightLight(params.enable),
            'enable_night_light': () => systemManager.toggleNightLight(true),
            'disable_night_light': () => systemManager.toggleNightLight(false),
            'night_light_on': () => systemManager.toggleNightLight(true),
            'night_light_off': () => systemManager.toggleNightLight(false),
            'empty_recycle_bin': () => systemManager.emptyRecycleBin(),
            'clear_recycle_bin': () => systemManager.emptyRecycleBin(),
            'delete_trash': () => systemManager.emptyRecycleBin(),
            'get_disk_space': () => systemManager.getDiskSpace(),
            'check_disk_space': () => systemManager.getDiskSpace(),
            'disk_usage': () => systemManager.getDiskSpace(),
            'set_volume': () => systemManager.setVolume(params.volume),
            'change_volume': () => systemManager.setVolume(params.volume),
            'adjust_volume': () => systemManager.setVolume(params.volume),
            'lock_computer': () => systemManager.lockComputer(),
            'lock_screen': () => systemManager.lockComputer(),
            'lock_pc': () => systemManager.lockComputer(),
            'sleep_computer': () => systemManager.sleep(),
            'sleep': () => systemManager.sleep(),
            'suspend': () => systemManager.sleep(),
            'get_network_status': () => systemManager.getNetworkStatus(),
            'network_status': () => systemManager.getNetworkStatus(),
            'network_info': () => systemManager.getNetworkStatus(),
            'run_disk_cleanup': () => systemManager.runDiskCleanup(),
            'disk_cleanup': () => systemManager.runDiskCleanup(),
            'cleanup_disk': () => systemManager.runDiskCleanup(),
            'toggle_windows_defender': () => systemManager.toggleWindowsDefender(params.enable),
            'enable_defender': () => systemManager.toggleWindowsDefender(true),
            'disable_defender': () => systemManager.toggleWindowsDefender(false),
            
            // Info
            'get_mouse_position': () => this.getMousePosition(),
            'get_screen_size': () => this.getScreenSize(),
            
            // Interactive choices
            'detect_browsers': () => this.detectInstalledBrowsers(),
            'detect_chrome_profiles': () => this.detectChromeProfiles(),
            'detect_edge_profiles': () => this.detectEdgeProfiles(),

            // ── OS Tasks (os-tasks.js) ──────────────────────────────────────
            // Win+R / run commands
            'run_winr': () => osTasks.runWinR(params.command),
            'open_run_dialog': () => osTasks.runWinR(params.command),
            'open_msconfig': () => osTasks.openMsconfig(),
            'open_services': () => osTasks.openServices(),
            'open_device_manager': () => osTasks.openDeviceManager(),
            'open_disk_management': () => osTasks.openDiskMgmt(),
            'open_regedit': () => osTasks.openRegedit(),
            'open_registry': () => osTasks.openRegedit(),
            'open_event_viewer': () => osTasks.openEventViewer(),
            'open_task_scheduler': () => osTasks.openTaskScheduler(),
            'open_group_policy': () => osTasks.openGroupPolicy(),
            'open_firewall': () => osTasks.openFirewall(),
            'open_network_connections': () => osTasks.openNetworkConns(),
            'open_power_options': () => osTasks.openPowerOptions(),
            'open_programs_features': () => osTasks.openProgramsAndFeatures(),
            'open_user_accounts': () => osTasks.openUserAccounts(),
            'open_display_settings': () => osTasks.openDisplaySettings(),
            'open_windows_update': () => osTasks.openWindowsUpdate(),
            'open_apps_settings': () => osTasks.openAppsSettings(),
            'open_bluetooth_settings': () => osTasks.openBluetooth(),
            'open_task_manager': () => osTasks.openTaskManager(),
            'open_control_panel': () => osTasks.openControlPanel(),
            'open_cmd': () => osTasks.openCmd(),
            'open_powershell': () => osTasks.openPowershell(),
            // Cache clearing
            'clear_cache': () => osTasks.clearCache(params.cache_type || params.type || 'all'),
            'clear_temp': () => osTasks.clearWindowsTemp(),
            'clear_windows_temp': () => osTasks.clearWindowsTemp(),
            'flush_dns': () => osTasks.flushDns(),
            'clear_dns_cache': () => osTasks.flushDns(),
            'clear_arp_cache': () => osTasks.clearArpCache(),
            'clear_icon_cache': () => osTasks.clearIconCache(),
            'clear_font_cache': () => osTasks.clearFontCache(),
            'clear_windows_update_cache': () => osTasks.clearWindowsUpdateCache(),
            'clear_chrome_cache': () => osTasks.clearChromeCache(),
            'clear_edge_cache': () => osTasks.clearEdgeCache(),
            'clear_firefox_cache': () => osTasks.clearFirefoxCache(),
            'clear_browser_cache': () => osTasks.clearBrowserCache(),
            'clear_all_cache': () => osTasks.clearCache('all'),
            // Network
            'reset_network': () => osTasks.resetNetwork(),
            'ping_host': () => osTasks.pingHost(params.host, params.count),
            'check_port': () => osTasks.checkPort(params.host, params.port),
            'get_public_ip': () => osTasks.getPublicIp(),
            'get_wifi_networks': () => osTasks.getWifiNetworks(),
            'connect_wifi': () => osTasks.connectWifi(params.ssid, params.password),
            // Services
            'manage_service': () => osTasks.manageService(params.service_name || params.name, params.action),
            'start_service': () => osTasks.manageService(params.service_name || params.name, 'start'),
            'stop_service': () => osTasks.manageService(params.service_name || params.name, 'stop'),
            'restart_service': () => osTasks.manageService(params.service_name || params.name, 'restart'),
            'list_services': () => osTasks.listServices(params.filter),
            // Processes
            'get_running_processes': () => osTasks.getRunningProcesses(params.sort_by),
            'list_processes': () => osTasks.getRunningProcesses(params.sort_by),
            'kill_process': () => osTasks.killProcess(params.name_or_pid || params.name || params.pid),
            'end_process': () => osTasks.killProcess(params.name_or_pid || params.name || params.pid),
            'get_process_details': () => osTasks.getProcessDetails(params.name),
            // Startup
            'list_startup_programs': () => osTasks.listStartupPrograms(),
            'toggle_startup_program': () => osTasks.toggleStartupProgram(params.name, params.enable !== false),
            'enable_startup_program': () => osTasks.toggleStartupProgram(params.name, true),
            'disable_startup_program': () => osTasks.toggleStartupProgram(params.name, false),
            // Windows Update
            'check_windows_update': () => osTasks.checkWindowsUpdate(),
            'windows_update': () => osTasks.checkWindowsUpdate(),
            'get_update_history': () => osTasks.getWindowsUpdateHistory(),
            // Event logs
            'get_event_logs': () => osTasks.getEventLogs(params.log_type, params.count, params.level),
            'clear_event_log': () => osTasks.clearEventLog(params.log_type || 'Application'),
            // Restore points
            'create_restore_point': () => osTasks.createRestorePoint(params.description || 'AI Created Restore Point'),
            'list_restore_points': () => osTasks.listRestorePoints(),
            // Installed apps
            'get_installed_apps': () => osTasks.getInstalledApps(),
            'list_installed_apps': () => osTasks.getInstalledApps(),
            'uninstall_app': () => osTasks.uninstallApp(params.app_name || params.name),
            // Environment variables
            'get_env_variable': () => osTasks.getEnvVariable(params.name),
            'set_env_variable': () => osTasks.setEnvVariable(params.name, params.value, params.scope),
            'delete_env_variable': () => osTasks.deleteEnvVariable(params.name, params.scope),
            'manage_env_variable': () => osTasks.setEnvVariable(params.name, params.value, params.scope),
            // Registry (user-safe only)
            'read_registry': () => osTasks.readRegistry(params.key_path, params.value_name),
            'write_registry': () => osTasks.writeRegistry(params.key_path, params.value_name, params.value, params.value_type),
            // Disk
            'get_disk_health': () => osTasks.getDiskHealth(),
            'disk_health': () => osTasks.getDiskHealth(),
            'analyze_storage': () => osTasks.analyzeStorageByFolder(params.path),
            'run_disk_cleanup_silent': () => osTasks.runDiskCleanupSilent(),
            // Power management
            'get_power_plan': () => osTasks.getPowerPlan(),
            'set_power_plan': () => osTasks.setPowerPlan(params.plan),
            'hibernate': () => osTasks.hibernate(),
            'hibernate_computer': () => osTasks.hibernate(),
            'restart_computer': () => osTasks.restart(params.delay || 0),
            'reboot': () => osTasks.restart(params.delay || 0),
            'shutdown_computer': () => osTasks.shutdown(params.delay || 0),
            'cancel_shutdown': () => osTasks.cancelShutdown(),
            // Security
            'windows_defender_scan': () => osTasks.quickScan(),
            'quick_scan': () => osTasks.quickScan(),
            'get_defender_status': () => osTasks.getDefenderStatus(),
            'check_firewall': () => osTasks.checkFirewallStatus(),
            // Scheduled tasks
            'list_scheduled_tasks': () => osTasks.listScheduledTasks(),
            'run_scheduled_task': () => osTasks.runScheduledTask(params.task_name || params.name),
            'disable_scheduled_task': () => osTasks.disableScheduledTask(params.task_name || params.name),
            // Display
            'toggle_dark_mode': () => osTasks.toggleDarkMode(params.enable !== false),
            'enable_dark_mode': () => osTasks.toggleDarkMode(true),
            'disable_dark_mode': () => osTasks.toggleDarkMode(false),
            'light_mode': () => osTasks.toggleDarkMode(false),
            'dark_mode': () => osTasks.toggleDarkMode(true),
            'set_taskbar_position': () => osTasks.setTaskbarPosition(params.position),
            'refresh_desktop': () => osTasks.refreshDesktop(),
            // Misc OS
            'get_clipboard': () => osTasks.getClipboard(),
            'set_clipboard': () => osTasks.setClipboard(params.text),
            'clipboard_get': () => osTasks.getClipboard(),
            'clipboard_set': () => osTasks.setClipboard(params.text),
            'show_notification': () => osTasks.showNotification(params.title, params.message || params.body, params.duration),
            'toast_notification': () => osTasks.showNotification(params.title, params.message || params.body, params.duration),
            'open_cmd_admin': () => osTasks.openCommandPromptAsAdmin(),
            'open_powershell_admin': () => osTasks.openPowerShellAsAdmin(),
            'get_system_health': () => osTasks.getFullSystemHealth(),
            'full_system_health': () => osTasks.getFullSystemHealth(),
            'list_fonts': () => osTasks.listInstalledFonts(),
            'get_installed_fonts': () => osTasks.listInstalledFonts(),

            // ── Browser Automation (browser-automation.js) ──────────────────
            'browser_open': () => browserAutomation.open(params.url, params.browser),
            'open_browser': () => browserAutomation.open(params.url, params.browser),
            'browser_navigate': () => browserAutomation.navigate(params.url),
            'navigate_to': () => browserAutomation.navigate(params.url),
            'browser_click': () => browserAutomation.click(params.selector || params.element),
            'browser_type': () => browserAutomation.type(params.selector || params.element, params.text),
            'browser_fill': () => browserAutomation.type(params.selector || params.element, params.text),
            'browser_get_text': () => browserAutomation.getText(params.selector),
            'browser_get_page_info': () => browserAutomation.getPageInfo(),
            'browser_screenshot': () => browserAutomation.screenshot(),
            'browser_scroll': () => browserAutomation.scroll(params.direction || 'down', params.amount || 300),
            'browser_scroll_up': () => browserAutomation.scroll('up', params.amount || 300),
            'browser_scroll_down': () => browserAutomation.scroll('down', params.amount || 300),
            'browser_wait_for': () => browserAutomation.waitFor(params.selector, params.timeout),
            'browser_close': () => browserAutomation.close(),
            'close_browser': () => browserAutomation.close(),
            'browser_login': () => browserAutomation.login(params.url, params.username, params.password, params.user_field, params.pass_field),
            'web_login': () => browserAutomation.login(params.url, params.username, params.password, params.user_field, params.pass_field),
            'browser_smart_login': () => browserAutomation.smartLogin(params.url, params.name, params.email || params.username, params.password, params.is_new_user || false),
            'browser_signup': () => browserAutomation.smartLogin(params.url, params.name, params.email, params.password, true),
            'browser_create_account': () => browserAutomation.smartLogin(params.url, params.name, params.email, params.password, true),
            'browser_search_in_page': () => browserAutomation.searchInPage(params.text || params.query),
            'browser_chat': () => browserAutomation.searchInPage(params.text || params.message || params.query),
            'browser_detect_page': () => browserAutomation.detectPageState(),
            'browser_auto_handle_blockers': () => browserAutomation.autoHandleBlockers(),
            'browser_page_screenshot_b64': async () => {
                const data = await browserAutomation.pageScreenshotBase64();
                return data ? { success: true, data } : { success: false, error: 'no page' };
            },
            'browser_search': () => browserAutomation.googleSearch(params.query),
            'google_search': () => browserAutomation.googleSearch(params.query),
            'web_search': () => browserAutomation.googleSearch(params.query),
            'browser_send_gmail': () => browserAutomation.sendGmail(params.to, params.subject, params.body, params.account_email || params.from || ''),
            'send_gmail': () => browserAutomation.sendGmail(params.to, params.subject, params.body, params.account_email || params.from || ''),
            'email_via_gmail': () => browserAutomation.sendGmail(params.to, params.subject, params.body, params.account_email || params.from || ''),
            'browser_open_gmail': () => browserAutomation.openGmailAccount(params.email || params.account_email || ''),
            'open_gmail': () => browserAutomation.openGmailAccount(params.email || params.account_email || ''),
            'browser_youtube_search': () => browserAutomation.youtubeSearch(params.query),
            'youtube_search': () => browserAutomation.youtubeSearch(params.query),
            'browser_go_back': () => browserAutomation.goBack(),
            'browser_go_forward': () => browserAutomation.goForward(),
            'browser_reload': () => browserAutomation.reload(),
            'browser_new_tab': () => browserAutomation.newTab(params.url),
            'browser_execute_script': () => browserAutomation.executeScript(params.script || params.js),
            'browser_fill_form': () => browserAutomation.fillForm(params.fields),
            'install_playwright': () => browserAutomation.installPlaywright(),
            'check_browser_ready': () => ({ success: true, ready: browserAutomation.isAvailable(), message: browserAutomation.isAvailable() ? 'Playwright ready' : 'Playwright not installed' })
        };

        const handler = actions[actionName];
        if (handler) {
            return await handler();
        } else {
            return { success: false, error: `Unknown action: ${actionName}` };
        }
    }

    // Focus an application
    async focusApp(appName) {
        const resolved = resolveAppName(appName);
        const processName = path.basename(resolved, '.exe').toUpperCase();
        
        const focused = await this.focusApplicationWindow(appName, processName);
        if (focused) {
            return { success: true, message: `Focused ${appName}` };
        } else {
            return { success: false, error: `Could not focus ${appName}` };
        }
    }

    // Wait/delay action
    async wait(ms) {
        await this.delay(ms);
        return { success: true, message: `Waited ${ms}ms` };
    }

    // ========================================
    // PowerPoint COM Automation Actions
    // ========================================

    /**
     * Apply design theme to PowerPoint presentation using keyboard shortcuts
     */
    async ppt_apply_theme(themeName) {
        try {
            // Use Design tab keyboard shortcut approach
            // Alt+G opens Design tab, then use arrow keys to select theme
            await this.wait(500);
            await this.pressKey('alt+g');  // Open Design tab
            await this.wait(300);
            await this.pressKey('h');  // Themes dropdown
            await this.wait(300);
            
            // Navigate to specific themes (Ion, Facet, etc.)
            const themeIndex = {
                'ion': 2,
                'facet': 3,
                'integral': 4,
                'ion_boardroom': 5,
                'office_theme': 1
            };
            
            const index = themeIndex[themeName.toLowerCase()] || 2;
            for (let i = 0; i < index; i++) {
                await this.pressKey('down');
                await this.wait(100);
            }
            
            await this.pressKey('enter');
            await this.wait(1000);
            
            return { 
                success: true, 
                message: `Applied theme: ${themeName}` 
            };
        } catch (error) {
            console.error('PowerPoint theme error:', error);
            return { 
                success: false, 
                error: `Failed to apply theme: ${error.message}` 
            };
        }
    }

    /**
     * Add animation to PowerPoint slide using keyboard shortcuts
     */
    async ppt_add_animation(animationType, applyToAll = false) {
        try {
            // Go to first slide
            await this.pressKey('home');
            await this.wait(500);
            
            // Select the title text box (click in center of slide)
            await this.pressKey('tab');
            await this.wait(300);
            
            // Open Animations tab with Alt+A
            await this.pressKey('alt+a');
            await this.wait(800);
            
            // Press F to get to Fade animation directly
            await this.pressKey('f');
            await this.wait(300);
            await this.pressKey('f');  // Press F again to select Fade
            await this.wait(500);
            
            return { 
                success: true, 
                message: `Added ${animationType} animation` 
            };
        } catch (error) {
            console.error('PowerPoint animation error:', error);
            return { 
                success: false, 
                error: `Failed to add animation: ${error.message}` 
            };
        }
    }

    /**
     * Change slide layout in PowerPoint
     */
    async ppt_change_layout(layoutName) {
        try {
            // Check if PowerPoint is active
            const isActive = await powerPointCOM.isPowerPointActive();
            if (!isActive) {
                return { 
                    success: false, 
                    error: 'No active PowerPoint presentation found. Please create slides first.' 
                };
            }

            // Change layout
            await powerPointCOM.changeLayout(layoutName);
            return { 
                success: true, 
                message: `Changed layout to: ${layoutName}` 
            };
        } catch (error) {
            console.error('PowerPoint layout error:', error);
            return { 
                success: false, 
                error: `Failed to change layout: ${error.message}` 
            };
        }
    }

    // ========================================
    // Paint Helper Methods (Cursor-Based)
    // ========================================

    /**
     * Draw line in Paint using cursor automation
     */
    async paintDrawLine(startX, startY, endX, endY) {
        try {
            // Draw line by dragging
            await this.drag(startX, startY, endX, endY);
            return { success: true, message: 'Line drawn in Paint' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Select tool in Paint (pencil, brush, line, etc.)
     * This would require vision-guided cursor automation
     */
    async paintSelectTool(toolName) {
        // This is a placeholder - would need cursor coordinates from vision AI
        return { 
            success: true, 
            message: `Paint tool selection: ${toolName} (requires vision guidance)` 
        };
    }

    /**
     * Select color in Paint
     */
    async paintSelectColor(color) {
        // Placeholder - would need vision-guided cursor automation
        return { 
            success: true, 
            message: `Paint color selection: ${color} (requires vision guidance)` 
        };
    }

    // ========================================
    // File System Utilities
    // ========================================

    /**
     * Check if file exists and return result with suggested alternative
     */
    async checkFileExists(filepath) {
        try {
            // Expand environment variables
            const expandedPath = filepath.replace(/%([^%]+)%/g, (_, key) => process.env[key] || '');
            
            const fs = require('fs').promises;
            try {
                await fs.access(expandedPath);
                // File exists - suggest numbered alternative
                const uniquePath = await this.findUniqueFilename(expandedPath);
                return {
                    success: true,
                    exists: true,
                    original_path: expandedPath,
                    suggested_path: uniquePath,
                    message: `File exists. Suggested alternative: ${path.basename(uniquePath)}`
                };
            } catch {
                // File doesn't exist - safe to use
                return {
                    success: true,
                    exists: false,
                    path: expandedPath,
                    message: 'File does not exist - safe to save'
                };
            }
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Generate unique filename by adding numbers if file exists
     */
    async generateUniqueFilename(filepath, contentDescription = '') {
        try {
            // If no filepath provided, generate from content description
            if (!filepath || filepath.trim() === '') {
                const timestamp = new Date().toISOString().split('T')[0].replace(/-/g, '_');
                const safeName = contentDescription
                    ? contentDescription.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30)
                    : 'Document';
                filepath = `C:\\Users\\${process.env.USERNAME}\\Documents\\${safeName}_${timestamp}.docx`;
            }

            // Expand environment variables
            const expandedPath = filepath.replace(/%([^%]+)%/g, (_, key) => process.env[key] || '');
            
            // Find unique filename
            const uniquePath = await this.findUniqueFilename(expandedPath);
            
            const wasNumbered = uniquePath !== expandedPath;
            return {
                success: true,
                path: uniquePath,
                original_path: expandedPath,
                was_numbered: wasNumbered,
                message: wasNumbered 
                    ? `Generated unique filename: ${path.basename(uniquePath)} (original name was taken)`
                    : `Filename is unique: ${path.basename(uniquePath)}`
            };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Helper: Find unique filename by adding _2, _3, etc.
     */
    async findUniqueFilename(filepath) {
        const fs = require('fs').promises;
        const dir = path.dirname(filepath);
        const ext = path.extname(filepath);
        const base = path.basename(filepath, ext);
        
        let counter = 2;
        let testPath = filepath;
        
        while (true) {
            try {
                await fs.access(testPath);
                // File exists, try next number
                testPath = path.join(dir, `${base}_${counter}${ext}`);
                counter++;
            } catch {
                // File doesn't exist - this is our unique name
                return testPath;
            }
        }
    }

    /**
     * Search for files in common locations
     */
    async searchFiles(searchTerm, fileType = '', searchLocation = '') {
        try {
            // Determine search locations
            const locations = searchLocation 
                ? [searchLocation.replace(/%([^%]+)%/g, (_, key) => process.env[key] || '')]
                : [
                    path.join(process.env.USERPROFILE, 'Documents'),
                    path.join(process.env.USERPROFILE, 'Desktop'),
                    path.join(process.env.USERPROFILE, 'Downloads')
                ];

            // Determine file extensions to search
            const extensions = fileType 
                ? this.getFileExtensions(fileType)
                : ['.docx', '.xlsx', '.pptx', '.pdf', '.txt', '.pub'];

            const command = `
                $searchTerm = "${searchTerm}"
                $locations = @(${locations.map(l => `"${l.replace(/\\/g, '\\\\')}"`).join(', ')})
                $extensions = @(${extensions.map(e => `"${e}"`).join(', ')})
                $results = @()
                
                foreach ($location in $locations) {
                    if (Test-Path $location) {
                        foreach ($ext in $extensions) {
                            $files = Get-ChildItem -Path $location -Filter "*$searchTerm*$ext" -File -ErrorAction SilentlyContinue
                            foreach ($file in $files) {
                                $results += [PSCustomObject]@{
                                    Name = $file.Name
                                    Path = $file.FullName
                                    Size = $file.Length
                                    Modified = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                                }
                            }
                        }
                    }
                }
                
                if ($results.Count -gt 0) {
                    $results | ConvertTo-Json -Compress
                } else {
                    Write-Output "NO_RESULTS"
                }
            `;

            return new Promise((resolve, reject) => {
                exec(`powershell -Command "${command.replace(/"/g, '\\"')}"`, { 
                    encoding: 'utf8',
                    maxBuffer: 10 * 1024 * 1024 
                }, (error, stdout, stderr) => {
                    if (error && !stdout.includes('NO_RESULTS')) {
                        return resolve({ 
                            success: false, 
                            error: error.message,
                            results: []
                        });
                    }

                    if (stdout.trim() === 'NO_RESULTS' || stdout.trim() === '') {
                        return resolve({
                            success: true,
                            results: [],
                            message: `No files found matching "${searchTerm}"`
                        });
                    }

                    try {
                        const results = JSON.parse(stdout.trim());
                        const fileList = Array.isArray(results) ? results : [results];
                        resolve({
                            success: true,
                            results: fileList,
                            count: fileList.length,
                            message: `Found ${fileList.length} file(s) matching "${searchTerm}"`
                        });
                    } catch (parseError) {
                        resolve({ 
                            success: false, 
                            error: 'Failed to parse results',
                            results: []
                        });
                    }
                });
            });
        } catch (error) {
            return { success: false, error: error.message, results: [] };
        }
    }

    /**
     * Get file extensions for file type
     */
    getFileExtensions(fileType) {
        const typeMap = {
            'word': ['.docx', '.doc'],
            'excel': ['.xlsx', '.xls'],
            'powerpoint': ['.pptx', '.ppt'],
            'pdf': ['.pdf'],
            'text': ['.txt'],
            'publisher': ['.pub'],
            'all': ['.docx', '.xlsx', '.pptx', '.pdf', '.txt', '.pub']
        };
        return typeMap[fileType.toLowerCase()] || ['.docx', '.xlsx', '.pptx'];
    }
}

module.exports = new ActionExecutor();

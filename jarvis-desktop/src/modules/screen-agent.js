// ============================================================
// Pecifics Screen Agent — Vision-Based Desktop Control
// ============================================================
// Instead of hardcoded CSS selectors, this module:
//   1. Takes a screenshot of the screen
//   2. Sends it to the AI vision model (Gemini)
//   3. Gets back pixel coordinates for where to click/type
//   4. Executes the action via OS-level mouse/keyboard
//
// This is the same approach used by Claude Computer Use,
// Google Mariner, and OpenAI Operator.
// ============================================================

const { exec } = require('child_process');
const os = require('os');

// ─────────────────────────────────────────────────────────────
// Win32 mouse/keyboard via PowerShell (works on entire desktop)
// ─────────────────────────────────────────────────────────────

const WIN32_TYPES = `
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ScreenAgent {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(int f, int dx, int dy, int d, int i);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
    public const int LEFTDOWN   = 0x02;
    public const int LEFTUP     = 0x04;
    public const int RIGHTDOWN  = 0x08;
    public const int RIGHTUP    = 0x10;
    public const int MIDDLEDOWN = 0x20;
    public const int MIDDLEUP   = 0x40;
    public const int WHEEL      = 0x0800;
}
[StructLayout(LayoutKind.Sequential)]
public struct POINT { public int X; public int Y; }
"@
`;

function ps(script, timeout = 8000) {
    return new Promise((resolve) => {
        const cmd = `powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "${script.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`;
        exec(cmd, { maxBuffer: 1024 * 512, timeout }, (err, stdout) => {
            resolve({ success: !err, output: (stdout || '').trim(), error: err ? err.message : null });
        });
    });
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─────────────────────────────────────────────────────────────
// Screen Agent Class
// ─────────────────────────────────────────────────────────────

class ScreenAgent {

    /**
     * Move mouse cursor to (x, y) screen coordinates.
     */
    async moveTo(x, y) {
        const r = await ps(`${WIN32_TYPES}; [ScreenAgent]::SetCursorPos(${Math.round(x)}, ${Math.round(y)})`);
        return { success: r.success, message: `Moved to (${x}, ${y})` };
    }

    /**
     * Click at screen coordinates (x, y).
     * @param {"left"|"right"|"middle"} button
     */
    async clickAt(x, y, button = 'left') {
        await this.moveTo(x, y);
        await delay(80);

        let downFlag, upFlag;
        if (button === 'right')       { downFlag = 'RIGHTDOWN';  upFlag = 'RIGHTUP'; }
        else if (button === 'middle') { downFlag = 'MIDDLEDOWN'; upFlag = 'MIDDLEUP'; }
        else                          { downFlag = 'LEFTDOWN';   upFlag = 'LEFTUP'; }

        const script = `${WIN32_TYPES}; [ScreenAgent]::mouse_event([ScreenAgent]::${downFlag}, 0, 0, 0, 0); Start-Sleep -Milliseconds 60; [ScreenAgent]::mouse_event([ScreenAgent]::${upFlag}, 0, 0, 0, 0)`;
        const r = await ps(script);
        return { success: r.success, message: `Clicked (${button}) at (${x}, ${y})` };
    }

    /**
     * Double-click at screen coordinates.
     */
    async doubleClickAt(x, y) {
        await this.clickAt(x, y, 'left');
        await delay(80);
        await this.clickAt(x, y, 'left');
        return { success: true, message: `Double-clicked at (${x}, ${y})` };
    }

    /**
     * Type text at the current cursor position using SendKeys.
     * For long or special text, uses clipboard paste for reliability.
     */
    async typeText(text) {
        if (!text) return { success: true, message: 'Nothing to type' };

        // Use clipboard for reliability (handles special chars, long text)
        const escaped = text.replace(/'/g, "''");
        const script = `
            Add-Type -AssemblyName System.Windows.Forms;
            [System.Windows.Forms.Clipboard]::SetText('${escaped}');
            Start-Sleep -Milliseconds 100;
            [System.Windows.Forms.SendKeys]::SendWait('^v');
        `;
        const r = await ps(script, 10000);
        return { success: r.success, message: `Typed ${text.length} chars` };
    }

    /**
     * Press a keyboard key (Enter, Tab, Escape, Backspace, etc.)
     * Also supports combos: "ctrl+a", "alt+f4", "ctrl+shift+t"
     */
    async pressKey(key) {
        if (!key) return { success: false, error: 'No key specified' };

        // Map common key names to SendKeys syntax
        const keyMap = {
            'enter':     '{ENTER}',
            'return':    '{ENTER}',
            'tab':       '{TAB}',
            'escape':    '{ESC}',
            'esc':       '{ESC}',
            'backspace': '{BACKSPACE}',
            'delete':    '{DELETE}',
            'space':     ' ',
            'up':        '{UP}',
            'down':      '{DOWN}',
            'left':      '{LEFT}',
            'right':     '{RIGHT}',
            'home':      '{HOME}',
            'end':       '{END}',
            'pageup':    '{PGUP}',
            'pagedown':  '{PGDN}',
            'f1': '{F1}', 'f2': '{F2}', 'f3': '{F3}', 'f4': '{F4}',
            'f5': '{F5}', 'f6': '{F6}', 'f7': '{F7}', 'f8': '{F8}',
            'f9': '{F9}', 'f10': '{F10}', 'f11': '{F11}', 'f12': '{F12}',
        };

        const lower = key.toLowerCase().trim();

        // Handle combos like ctrl+a, ctrl+shift+t
        if (lower.includes('+')) {
            const parts = lower.split('+');
            let prefix = '';
            let mainKey = parts[parts.length - 1];
            for (let i = 0; i < parts.length - 1; i++) {
                const mod = parts[i].trim();
                if (mod === 'ctrl' || mod === 'control') prefix += '^';
                else if (mod === 'alt')                   prefix += '%';
                else if (mod === 'shift')                 prefix += '+';
                else if (mod === 'win')                   prefix += '^{ESC}'; // approximation
            }
            const mapped = keyMap[mainKey] || mainKey;
            const sendStr = `${prefix}${mapped}`;
            const script = `Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('${sendStr.replace(/'/g, "''")}')`;
            const r = await ps(script);
            return { success: r.success, message: `Pressed ${key}` };
        }

        const mapped = keyMap[lower] || key;
        const script = `Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('${mapped.replace(/'/g, "''")}')`;
        const r = await ps(script);
        return { success: r.success, message: `Pressed ${key}` };
    }

    /**
     * Scroll at current cursor position.
     * @param {"up"|"down"} direction
     * @param {number} clicks — number of scroll wheel clicks (default 3)
     */
    async scroll(direction = 'down', clicks = 3) {
        const value = direction === 'up' ? (clicks * 120) : -(clicks * 120);
        const script = `${WIN32_TYPES}; [ScreenAgent]::mouse_event([ScreenAgent]::WHEEL, 0, 0, ${value}, 0)`;
        const r = await ps(script);
        return { success: r.success, message: `Scrolled ${direction} ${clicks} clicks` };
    }

    /**
     * Execute a vision action returned by the /vision_act endpoint.
     * @param {Object} act — {action, x, y, button, text, key, direction, clicks, description}
     * @returns {Object} — {success, message}
     */
    async executeVisionAction(act) {
        if (!act || !act.action) {
            return { success: false, error: 'No action provided' };
        }

        const action = act.action.toLowerCase();

        switch (action) {
            case 'click':
                return await this.clickAt(act.x, act.y, act.button || 'left');

            case 'double_click':
            case 'doubleclick':
                return await this.doubleClickAt(act.x, act.y);

            case 'right_click':
            case 'rightclick':
                return await this.clickAt(act.x, act.y, 'right');

            case 'type':
                // If coordinates given, click there first then type
                if (act.x !== undefined && act.y !== undefined) {
                    await this.clickAt(act.x, act.y);
                    await delay(200);
                }
                return await this.typeText(act.text || '');

            case 'key':
            case 'press_key':
            case 'hotkey':
                return await this.pressKey(act.key || '');

            case 'scroll':
                // If coordinates given, move there first
                if (act.x !== undefined && act.y !== undefined) {
                    await this.moveTo(act.x, act.y);
                    await delay(100);
                }
                return await this.scroll(act.direction || 'down', act.clicks || 3);

            case 'move':
                return await this.moveTo(act.x, act.y);

            case 'drag':
                return await this._drag(act.x, act.y, act.end_x, act.end_y);

            case 'wait':
                await delay(act.duration || 1000);
                return { success: true, message: `Waited ${act.duration || 1000}ms` };

            case 'done':
                return { success: true, done: true, message: act.description || 'Task complete' };

            case 'fail':
                return { success: false, done: true, error: act.description || 'Cannot proceed' };

            default:
                return { success: false, error: `Unknown vision action: ${action}` };
        }
    }

    /**
     * Drag from (x1,y1) to (x2,y2).
     */
    async _drag(x1, y1, x2, y2) {
        await this.moveTo(x1, y1);
        await delay(100);
        await ps(`${WIN32_TYPES}; [ScreenAgent]::mouse_event([ScreenAgent]::LEFTDOWN, 0, 0, 0, 0)`);
        await delay(50);

        // Smooth drag in steps
        const steps = 15;
        for (let i = 1; i <= steps; i++) {
            const cx = Math.round(x1 + (x2 - x1) * (i / steps));
            const cy = Math.round(y1 + (y2 - y1) * (i / steps));
            await this.moveTo(cx, cy);
            await delay(15);
        }

        await delay(50);
        await ps(`${WIN32_TYPES}; [ScreenAgent]::mouse_event([ScreenAgent]::LEFTUP, 0, 0, 0, 0)`);
        return { success: true, message: `Dragged from (${x1},${y1}) to (${x2},${y2})` };
    }
}

module.exports = new ScreenAgent();

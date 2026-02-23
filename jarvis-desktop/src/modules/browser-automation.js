// ============================================================
// Pecifics Browser Automation Module
// Uses Playwright (Node.js) for reliable browser control
// Falls back to keyboard/mouse Control if Playwright unavailable
// ============================================================

const { exec, spawn } = require('child_process');
const path = require('path');
const os   = require('os');

let playwright = null;
let chromium   = null;
let browser    = null;
let page       = null;
let playwrightAvailable = false;

// Try to load Playwright at startup
(async () => {
    try {
        const pw = require('playwright');
        playwright = pw;
        chromium   = pw.chromium;
        playwrightAvailable = true;
        console.log('✅ Playwright loaded – rich browser automation available');
    } catch (e) {
        console.warn('⚠️  Playwright not installed. Browser automation will use keyboard/mouse fallback.');
        console.warn('    To enable: cd jarvis-desktop && npm install playwright && npx playwright install chromium');
    }
})();

// ─────────────────────────────────────────────────────────────
// PowerShell helper
// ─────────────────────────────────────────────────────────────
function ps(script) {
    return new Promise((resolve) => {
        exec(
            `powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "${script.replace(/"/g, '\\"').replace(/\n/g, ' ')}"`,
            { maxBuffer: 1024 * 1024 * 5, timeout: 30000 },
            (err, stdout, stderr) => {
                const out = (stdout || '').trim();
                resolve({ success: !err || !!out, output: out, error: err ? err.message : null });
            }
        );
    });
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─────────────────────────────────────────────────────────────
// Playwright session management
// ─────────────────────────────────────────────────────────────

async function ensureBrowser(browserName = 'chromium') {
    if (!playwrightAvailable) throw new Error('Playwright not available');
    if (!browser || !browser.isConnected()) {
        const launchOpts = { headless: false, args: ['--start-maximized'] };
        if (browserName === 'firefox') {
            browser = await playwright.firefox.launch(launchOpts);
        } else if (browserName === 'webkit' || browserName === 'safari') {
            browser = await playwright.webkit.launch(launchOpts);
        } else {
            // chromium – try to use installed Chrome/Edge first
            try {
                const { executablePath } = require('playwright');
                browser = await chromium.launch({
                    ...launchOpts,
                    channel: 'chrome',  // use installed Chrome
                });
            } catch {
                browser = await chromium.launch(launchOpts);
            }
        }
    }
    if (!page || page.isClosed()) {
        const ctx = await browser.newContext({ viewport: null }); // full screen
        page = await ctx.newPage();
    }
    return { browser, page };
}

async function getOrCreatePage() {
    if (!playwrightAvailable) return null;
    try {
        return (await ensureBrowser()).page;
    } catch (e) {
        console.error('Browser session error:', e.message);
        return null;
    }
}

// ─────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────

class BrowserAutomation {

    /**
     * Open a URL in a specific browser.
     * Uses Playwright if available, otherwise falls back to shell open.
     */
    async open(url, browserName = 'chrome') {
        if (playwrightAvailable) {
            try {
                const { page: p } = await ensureBrowser(browserName);
                await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
                const title = await p.title();
                return { success: true, message: `Opened: ${url}`, title };
            } catch (e) {
                console.warn('Playwright open failed, using shell fallback:', e.message);
            }
        }
        // Fallback: open with OS default
        return new Promise((resolve) => {
            exec(`start "" "${url}"`, (err) => {
                resolve({ success: !err, message: err ? err.message : `Opened ${url} in browser` });
            });
        });
    }

    /**
     * Navigate the current tab to a URL.
     */
    async navigate(url) {
        if (!url.startsWith('http')) url = 'https://' + url;
        const p = await getOrCreatePage();
        if (p) {
            try {
                await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
                const title = await p.title();
                return { success: true, message: `Navigated to: ${url}`, title };
            } catch (e) {
                return { success: false, error: e.message };
            }
        }
        // Fallback
        return this.open(url);
    }

    /**
     * Click an element by CSS selector, text content, or aria-label.
     * Falls back to JS-based element discovery when static selectors fail.
     */
    async click(selector) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            // Strategy 1: exact and common selector variants
            const strategies = [
                () => p.click(selector, { timeout: 4000 }),
                () => p.click(`text=${selector}`, { timeout: 4000 }),
                () => p.click(`[aria-label="${selector}"]`, { timeout: 4000 }),
                () => p.click(`[placeholder="${selector}"]`, { timeout: 4000 }),
                () => p.click(`[name="${selector}"]`, { timeout: 4000 }),
                () => p.click(`button:has-text("${selector}")`, { timeout: 4000 }),
                () => p.click(`a:has-text("${selector}")`, { timeout: 4000 }),
                () => p.click(`[data-testid="${selector}"]`, { timeout: 4000 }),
                () => p.click(`[id*="${selector.replace(/^[#.]/, '')}"]`, { timeout: 4000 }),
                () => p.click(`[class*="${selector.replace(/^[#.]/, '')}"]`, { timeout: 4000 }),
            ];
            for (const strat of strategies) {
                try { await strat(); return { success: true, message: `Clicked: ${selector}` }; }
                catch {}
            }
            // Strategy 2: JS-based smart finder — search all clickable elements for matching text/attr
            const hint = selector.replace(/^[#.\[\]]/, '').toLowerCase();
            const clicked = await p.evaluate((hint) => {
                const candidates = Array.from(document.querySelectorAll(
                    'button, a, [role="button"], input[type="submit"], input[type="button"], [onclick], [tabindex]'
                ));
                for (const el of candidates) {
                    const text = (el.innerText || el.value || el.getAttribute('aria-label') || el.getAttribute('title') || '').toLowerCase();
                    const id   = (el.id || '').toLowerCase();
                    const cls  = (el.className || '').toLowerCase();
                    if (text.includes(hint) || id.includes(hint) || cls.includes(hint)) {
                        el.scrollIntoView({ behavior: 'instant', block: 'center' });
                        el.click();
                        return true;
                    }
                }
                return false;
            }, hint);
            if (clicked) return { success: true, message: `Clicked element matching: ${selector}` };
            return { success: false, error: `Could not find element: ${selector}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Type text into a field by selector.
     * Falls back to JS-based input discovery when selector doesn't match.
     */
    async type(selector, text, clearFirst = true) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const hint = selector.toLowerCase();
            const isPassword = hint.includes('pass');
            const isEmail    = hint.includes('email') || hint.includes('user');

            const strategies = [
                selector,
                `[placeholder*="${selector.replace(/^[#.]/, '')}" i]`,
                `[name*="${selector.replace(/^[#.]/, '')}" i]`,
                `[id*="${selector.replace(/^[#.]/, '')}" i]`,
                `[aria-label*="${selector.replace(/^[#.]/, '')}" i]`,
                `input[type="${selector.replace(/^[#.]/, '')}"]`,
            ];
            if (isPassword) strategies.push('input[type="password"]');
            if (isEmail)    strategies.push('input[type="email"]', 'input[type="text"]');

            for (const sel of strategies) {
                try {
                    await p.waitForSelector(sel, { timeout: 4000, state: 'visible' });
                    if (clearFirst) await p.fill(sel, '');
                    await p.fill(sel, text);
                    return { success: true, message: `Typed into: ${selector}` };
                } catch {}
            }

            // JS-based smart finder — look for any visible input matching the hint
            const typed = await p.evaluate(({ hint, text, isPassword, isEmail }) => {
                const inputs = Array.from(document.querySelectorAll('input, textarea'));
                for (const el of inputs) {
                    if (el.offsetParent === null) continue; // hidden
                    const type  = (el.type || '').toLowerCase();
                    const id    = (el.id || '').toLowerCase();
                    const name  = (el.name || '').toLowerCase();
                    const ph    = (el.placeholder || '').toLowerCase();
                    const label = (el.getAttribute('aria-label') || '').toLowerCase();
                    const combined = id + name + ph + label + type;
                    if (
                        combined.includes(hint) ||
                        (isPassword && type === 'password') ||
                        (isEmail && (type === 'email' || combined.includes('email') || combined.includes('user')))
                    ) {
                        el.focus();
                        el.value = text;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        return true;
                    }
                }
                // Last resort: first visible text/email input
                for (const el of inputs) {
                    if (el.offsetParent === null) continue;
                    if (['text', 'email', 'password', 'search', ''].includes(el.type)) {
                        el.focus();
                        el.value = text;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        return true;
                    }
                }
                return false;
            }, { hint: selector.replace(/^[#.]/, '').toLowerCase(), text, isPassword, isEmail });

            if (typed) return { success: true, message: `Typed into field matching: ${selector}` };
            return { success: false, error: `Could not find input: ${selector}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Get text content of an element.
     */
    async getText(selector) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const text = await p.textContent(selector, { timeout: 5000 });
            return { success: true, text, message: text };
        } catch (e) {
            // Try inner text of body as fallback
            try {
                const bodyText = await p.evaluate(() => document.body.innerText);
                return { success: true, text: bodyText.substring(0, 2000), message: 'Page text (full body)' };
            } catch {
                return { success: false, error: e.message };
            }
        }
    }

    /**
     * Get the current page title and URL.
     */
    async getPageInfo() {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const title = await p.title();
            const url   = p.url();
            return { success: true, title, url };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Search Google and return first N result URLs + titles.
     */
    async googleSearch(query, resultCount = 5) {
        const p = await getOrCreatePage();
        if (!p) {
            // Fallback: just open Google in default browser
            const searchUrl = `https://www.google.com/search?q=${encodeURIComponent(query)}`;
            exec(`start "" "${searchUrl}"`);
            return { success: true, message: `Opened Google search for: ${query}`, results: [] };
        }
        try {
            await p.goto(`https://www.google.com/search?q=${encodeURIComponent(query)}&hl=en`, { waitUntil: 'domcontentloaded' });
            const results = await p.evaluate((count) => {
                const items = [];
                document.querySelectorAll('h3').forEach((h3) => {
                    const a = h3.closest('a');
                    if (a && items.length < count) {
                        items.push({ title: h3.innerText, url: a.href });
                    }
                });
                return items;
            }, resultCount);
            return { success: true, query, results, message: `Found ${results.length} results for: ${query}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Auto-fill and submit a login form.
     */
    async login(url, username, password, usernameSelector = '', passwordSelector = '') {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };

        try {
            await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 20000 });
            await delay(1500);

            // Detect username field
            const userSelectors = [
                usernameSelector,
                'input[type="email"]',
                'input[name="email"]',
                'input[name="username"]',
                'input[name="user"]',
                'input[id*="email"]',
                'input[id*="user"]',
                'input[placeholder*="email" i]',
                'input[placeholder*="username" i]',
                'input[type="text"]:first-of-type',
            ].filter(Boolean);

            const passSelectors = [
                passwordSelector,
                'input[type="password"]',
                'input[name="password"]',
                'input[name="pass"]',
                'input[id*="password"]',
                'input[id*="pass"]',
            ].filter(Boolean);

            let userFilled = false;
            for (const sel of userSelectors) {
                try {
                    await p.fill(sel, username, { timeout: 3000 });
                    userFilled = true;
                    break;
                } catch {}
            }
            if (!userFilled) return { success: false, error: 'Could not find username/email field' };

            await delay(500);

            // Some sites show password on next page
            try { await p.press(userSelectors[0], 'Tab'); } catch {}
            await delay(300);
            try { await p.press(userSelectors[0], 'Enter'); } catch {}
            await delay(1500);

            let passFilled = false;
            for (const sel of passSelectors) {
                try {
                    await p.fill(sel, password, { timeout: 3000 });
                    passFilled = true;
                    break;
                } catch {}
            }
            if (!passFilled) return { success: false, error: 'Could not find password field' };

            await delay(300);

            // Submit
            const submitSelectors = [
                'button[type="submit"]',
                'input[type="submit"]',
                'button:has-text("Sign in")',
                'button:has-text("Log in")',
                'button:has-text("Login")',
                'button:has-text("Continue")',
                '[data-testid*="submit"]',
            ];
            let submitted = false;
            for (const sel of submitSelectors) {
                try {
                    await p.click(sel, { timeout: 3000 });
                    submitted = true;
                    break;
                } catch {}
            }
            if (!submitted) {
                await p.keyboard.press('Enter');
            }

            await delay(3000); // Wait for navigation
            const title = await p.title();
            const urlNow = p.url();

            return {
                success: true,
                message: `Login attempted at ${url}`,
                currentPage: { title, url: urlNow },
            };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Detect the current page state: 'login' | 'signup' | 'chat' | 'search' | 'main'
     */
    async detectPageState() {
        const p = await getOrCreatePage();
        if (!p) return { success: false, state: 'unknown' };
        try {
            const state = await p.evaluate(() => {
                const body = (document.body && document.body.innerText || '').toLowerCase();
                const inputs = Array.from(document.querySelectorAll('input'));
                const hasPassword = inputs.some(i => i.type === 'password');
                const passCount   = inputs.filter(i => i.type === 'password').length;
                const hasNameField = inputs.some(i => {
                    const combined = (i.name + i.id + i.placeholder).toLowerCase();
                    return ['name','full','first','last'].some(k => combined.includes(k));
                });
                const hasSignupText = /create[\s\w]*account|sign[\s-]?up|register|get started/i.test(body);
                const hasLoginText  = /sign[\s-]?in|log[\s-]?in|welcome back|password/i.test(body);
                const hasChatInput  = !!(document.querySelector('#prompt-textarea') ||
                    document.querySelector('textarea[data-id]') ||
                    document.querySelector('div[contenteditable="true"][class*="prompt"]'));
                if (passCount >= 2 || (hasSignupText && hasNameField)) return 'signup';
                if (hasPassword) return 'login';
                if (hasChatInput) return 'chat';
                if (document.querySelector('input[type="search"], input[name="q"]')) return 'search';
                return 'main';
            });
            return { success: true, state, url: p.url(), title: await p.title() };
        } catch (e) {
            return { success: false, state: 'unknown', error: e.message };
        }
    }

    /**
     * Smart login — auto-detects login vs signup, handles multi-step flows.
     * isNewUser=true  → navigate to signup/register form and create account
     * isNewUser=false → sign into existing account
     */
    async smartLogin(url, name, email, password, isNewUser = false) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
            await delay(2000);

            const stateInfo = await this.detectPageState();
            let state = stateInfo.state;

            // ── New user: navigate to signup form ──
            if (isNewUser && state !== 'signup') {
                const signupLinks = [
                    'a:has-text("Sign up")', 'a:has-text("Create account")',
                    'a:has-text("Register")', 'button:has-text("Sign up")',
                    'a[href*="signup"]', 'a[href*="register"]',
                    'a[href*="join"]', 'a[href*="create"]',
                ];
                for (const sel of signupLinks) {
                    try { await p.click(sel, { timeout: 3000 }); await delay(2000); break; } catch {}
                }
                state = (await this.detectPageState()).state;
            }

            if (state === 'signup' || isNewUser) {
                return await this._doSignup(p, name, email || name, password);
            }
            return await this._doLoginSteps(p, email || name, password);
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /** Internal: fill signup form (name, email, passwords) and submit */
    async _doSignup(p, name, email, password) {
        try {
            // Name field
            if (name) {
                const nameSelectors = [
                    'input[name*="name" i]', 'input[id*="name" i]',
                    'input[placeholder*="name" i]', 'input[autocomplete="name"]',
                ];
                for (const sel of nameSelectors) {
                    try { await p.fill(sel, name, { timeout: 3000 }); break; } catch {}
                }
                await delay(300);
            }
            // Email
            const emailSelectors = [
                'input[type="email"]', 'input[name*="email" i]',
                'input[id*="email" i]', 'input[placeholder*="email" i]',
            ];
            for (const sel of emailSelectors) {
                try { await p.fill(sel, email, { timeout: 3000 }); break; } catch {}
            }
            await delay(300);
            // All password fields (new + confirm)
            const passFields = await p.$$('input[type="password"]');
            for (const field of passFields) {
                try { await field.fill(password); await delay(200); } catch {}
            }
            await delay(400);
            // Submit
            const submitSelectors = [
                'button[type="submit"]', 'input[type="submit"]',
                'button:has-text("Sign up")', 'button:has-text("Create account")',
                'button:has-text("Register")', 'button:has-text("Get started")',
                'button:has-text("Continue")', 'button:has-text("Next")',
            ];
            for (const sel of submitSelectors) {
                try { await p.click(sel, { timeout: 3000 }); break; } catch {}
            }
            await delay(3000);
            return { success: true, message: `Account creation attempted for ${email}`, action: 'signup' };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /** Internal: multi-step login (email → Enter → password may appear on next step) */
    async _doLoginSteps(p, usernameOrEmail, password) {
        try {
            const userSelectors = [
                'input[type="email"]', 'input[name="email"]', 'input[name="username"]',
                'input[name="user"]', 'input[id*="email" i]', 'input[id*="user" i]',
                'input[placeholder*="email" i]', 'input[placeholder*="username" i]',
                'input[autocomplete="email"]', 'input[autocomplete="username"]',
            ];
            let userFilled = false;
            let usedSel = userSelectors[0];
            for (const sel of userSelectors) {
                try {
                    await p.fill(sel, usernameOrEmail, { timeout: 3000 });
                    userFilled = true; usedSel = sel; break;
                } catch {}
            }
            if (!userFilled) return { success: false, error: 'Could not find email/username field' };
            await delay(400);

            // Try Continue/Next button first (multi-step sites like Google, Microsoft)
            const continueSelectors = [
                'button:has-text("Next")', 'button:has-text("Continue")',
                'input[value*="Next" i]', 'button[id*="next" i]',
            ];
            let stepped = false;
            for (const sel of continueSelectors) {
                try { await p.click(sel, { timeout: 2500 }); stepped = true; await delay(2000); break; } catch {}
            }
            if (!stepped) {
                // Press Enter to advance
                try { await p.press(usedSel, 'Enter'); await delay(1800); } catch {}
            }

            // Now fill password (may be on new step/page)
            const passSelectors = [
                'input[type="password"]', 'input[name="password"]',
                'input[name="pass"]', 'input[id*="password" i]',
            ];
            let passFilled = false;
            for (const sel of passSelectors) {
                try {
                    await p.fill(sel, password, { timeout: 4000 });
                    passFilled = true; break;
                } catch {}
            }
            if (!passFilled) return { success: false, error: 'Could not find password field' };
            await delay(300);

            // Submit
            const submitSelectors = [
                'button[type="submit"]', 'input[type="submit"]',
                'button:has-text("Sign in")', 'button:has-text("Log in")',
                'button:has-text("Login")', 'button:has-text("Continue")',
                'button:has-text("Next")', '[data-testid*="submit"]',
            ];
            let submitted = false;
            for (const sel of submitSelectors) {
                try { await p.click(sel, { timeout: 3000 }); submitted = true; break; } catch {}
            }
            if (!submitted) await p.keyboard.press('Enter');

            await delay(3500);
            const title  = await p.title();
            const urlNow = p.url();
            return { success: true, message: `Signed in successfully`, currentPage: { title, url: urlNow }, action: 'login' };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Type into the main search box or chat prompt on the current page and submit.
     * Works with ChatGPT, YouTube search, Google, Bing, site search bars, etc.
     */
    async searchInPage(text) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const searchSelectors = [
                // ChatGPT / AI chat prompts
                '#prompt-textarea',
                'textarea[data-id="root"]',
                'div[id="prompt-textarea"][contenteditable]',
                'p[data-placeholder]',
                // Generic textareas
                'textarea[placeholder*="message" i]',
                'textarea[placeholder*="ask" i]',
                'textarea[placeholder*="search" i]',
                'textarea[placeholder*="type" i]',
                // Standard search inputs
                'input[type="search"]',
                'input[name="q"]',
                'input[name="search_query"]',
                'input[placeholder*="search" i]',
                'input[placeholder*="ask" i]',
                // Contenteditable
                '[contenteditable="true"]',
            ];

            for (const sel of searchSelectors) {
                try {
                    await p.waitForSelector(sel, { timeout: 3000, state: 'visible' });
                    await p.click(sel);
                    await delay(300);
                    // contenteditable? use keyboard type
                    const isEditable = await p.$eval(sel, el => el.contentEditable === 'true').catch(() => false);
                    if (isEditable) {
                        // Clear existing content
                        await p.keyboard.press('Control+a');
                        await p.keyboard.press('Delete');
                        await p.keyboard.type(text, { delay: 30 });
                    } else {
                        await p.fill(sel, text);
                    }
                    await delay(400);
                    await p.keyboard.press('Enter');
                    return { success: true, message: `Sent: "${text}"` };
                } catch {}
            }
            return { success: false, error: 'No search or chat input found on this page' };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Scroll in the page.
     */
    async scroll(direction = 'down', amount = 500) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const delta = direction === 'up' ? -amount : amount;
            await p.mouse.wheel(0, delta);
            return { success: true, message: `Scrolled ${direction}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Wait for a selector to appear.
     */
    async waitFor(selector, timeout = 10000) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            await p.waitForSelector(selector, { timeout });
            return { success: true, message: `Element appeared: ${selector}` };
        } catch (e) {
            return { success: false, error: `Timeout waiting for: ${selector}` };
        }
    }

    /**
     * Take a screenshot of the current browser page as base64.
     */
    async pageScreenshotBase64() {
        const p = await getOrCreatePage();
        if (!p) return null;
        try {
            const buf = await p.screenshot({ type: 'jpeg', quality: 70, fullPage: false });
            return buf.toString('base64');
        } catch { return null; }
    }

    /**
     * DOM-based fast blocker detection and auto-handling.
     * Handles: cookie banners, GDPR popups, generic confirm/close dialogs.
     * Returns { handled: bool, what: string }
     */
    async autoHandleBlockers() {
        const p = await getOrCreatePage();
        if (!p) return { handled: false, what: 'no browser' };
        try {
            const acceptSelectors = [
                // Cookie consent
                'button:has-text("Accept all")', 'button:has-text("Accept All")',
                'button:has-text("Accept cookies")', 'button:has-text("Accept Cookies")',
                'button:has-text("Allow all")', 'button:has-text("Allow All")',
                'button:has-text("Allow cookies")', 'button:has-text("I Accept")',
                'button:has-text("I agree")', 'button:has-text("Agree")',
                '#onetrust-accept-btn-handler',
                'button[id*="accept" i][id*="cookie" i]',
                'button[class*="accept" i][class*="cookie" i]',
                // Confirm/dismiss
                'button:has-text("Got it")', 'button:has-text("OK")',
                'button:has-text("Close")', 'button:has-text("Dismiss")',
                'button:has-text("Continue")', 'button:has-text("Not now")',
                '[aria-label="Close"]', '[aria-label="Dismiss"]',
            ];
            for (const sel of acceptSelectors) {
                try {
                    const el = await p.$(sel);
                    if (el && await el.isVisible()) {
                        await el.click();
                        await delay(800);
                        return { handled: true, what: `Clicked: ${sel}` };
                    }
                } catch {}
            }
            return { handled: false, what: 'no known blocker found' };
        } catch (e) {
            return { handled: false, what: e.message };
        }
    }

    /**
     * Take a screenshot of the current browser page.
     */
    async screenshot() {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const imgPath = path.join(os.tmpdir(), `browser_shot_${Date.now()}.png`);
            await p.screenshot({ path: imgPath, fullPage: false });
            return { success: true, path: imgPath, message: `Screenshot saved: ${imgPath}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Execute JavaScript in the page context.
     */
    async executeScript(script) {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        try {
            const result = await p.evaluate(new Function(`return (${script})`));
            return { success: true, result: JSON.stringify(result), message: String(result) };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Fill a form with key-value pairs.
     */
    async fillForm(fields) {
        // fields = [{ selector: "...", value: "..." }, ...]
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open' };
        const results = [];
        for (const { selector, value } of fields) {
            try {
                await p.fill(selector, value, { timeout: 5000 });
                results.push({ selector, success: true });
            } catch (e) {
                results.push({ selector, success: false, error: e.message });
            }
        }
        const ok = results.filter(r => r.success).length;
        return { success: ok > 0, message: `Filled ${ok}/${results.length} form fields`, details: results };
    }

    /**
     * Go back/forward in browser history.
     */
    async goBack()    { const p = await getOrCreatePage(); if (p) { await p.goBack();    return { success: true, message: 'Went back' }; } return { success: false }; }
    async goForward() { const p = await getOrCreatePage(); if (p) { await p.goForward(); return { success: true, message: 'Went forward' }; } return { success: false }; }
    async reload()    { const p = await getOrCreatePage(); if (p) { await p.reload();    return { success: true, message: 'Page reloaded' }; } return { success: false }; }

    /**
     * Open a new tab.
     */
    async newTab(url = '') {
        if (!browser || !browser.isConnected()) return { success: false, error: 'Browser not open' };
        try {
            const ctx = browser.contexts()[0];
            const newPage = await ctx.newPage();
            if (url) await newPage.goto(url, { waitUntil: 'domcontentloaded' });
            page = newPage;
            return { success: true, message: `New tab opened${url ? ': ' + url : ''}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Close the browser completely.
     */
    async close() {
        try {
            if (browser) { await browser.close(); browser = null; page = null; }
            return { success: true, message: 'Browser closed' };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Resolve which Gmail account slot (u/0, u/1, ...) belongs to the given email.
     * Returns the index (0-based) or -1 if not found.
     */
    async resolveGmailIndex(email) {
        if (!playwrightAvailable) return -1;
        const emailLower = email.toLowerCase().trim();
        // Try up to 5 account slots
        for (let i = 0; i < 5; i++) {
            try {
                const { page: p } = await ensureBrowser();
                await p.goto(`https://mail.google.com/mail/u/${i}/`, { waitUntil: 'domcontentloaded', timeout: 12000 });
                await delay(1500);
                const url = p.url();
                // Redirected to login/accounts page means this slot isn't logged in
                if (url.includes('accounts.google.com')) break;
                // Look for the signed-in email shown in page DOM
                const pageEmail = await p.evaluate(() => {
                    // Gmail exposes the account email in several places
                    const el =
                        document.querySelector('a[aria-label*="@"]') ||
                        document.querySelector('div[data-email]') ||
                        document.querySelector('[aria-label*="Google Account"]') ||
                        document.querySelector('span.gb_mb') ||
                        document.querySelector('div.gb_Cb');
                    if (el) {
                        const label = el.getAttribute('aria-label') || el.getAttribute('data-email') || el.innerText || '';
                        const match = label.match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/i);
                        return match ? match[0].toLowerCase() : '';
                    }
                    // Fallback: scan all text nodes
                    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
                    while (walker.nextNode()) {
                        const t = walker.currentNode.nodeValue || '';
                        const m = t.match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/i);
                        if (m) return m[0].toLowerCase();
                    }
                    return '';
                });
                if (pageEmail && pageEmail === emailLower) return i;
            } catch { break; }
        }
        return -1;
    }

    /**
     * Open Gmail for a specific email account.
     * - If the account is already signed in, navigates to its slot.
     * - If not found, navigates to account chooser / prompts login.
     */
    async openGmailAccount(email) {
        try {
            const idx = email ? await this.resolveGmailIndex(email) : 0;
            if (idx >= 0) {
                const { page: p } = await ensureBrowser();
                await p.goto(`https://mail.google.com/mail/u/${idx}/`, { waitUntil: 'domcontentloaded', timeout: 20000 });
                return { success: true, message: `Opened Gmail for ${email} (account slot ${idx})`, accountIndex: idx };
            }
            // Account not found — navigate to account chooser so user can sign in or pick
            const { page: p } = await ensureBrowser();
            const chooserUrl = email
                ? `https://accounts.google.com/AccountChooser?Email=${encodeURIComponent(email)}&continue=https://mail.google.com`
                : 'https://mail.google.com';
            await p.goto(chooserUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
            return {
                success: true,
                message: `Account ${email} not found in signed-in accounts. Opened account chooser — please sign in.`,
                needsLogin: true
            };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Send an email via Gmail web interface (Playwright).
     * If accountEmail is given, switches to that Gmail account first.
     */
    async sendGmail(to, subject, body, accountEmail = '') {
        const p = await getOrCreatePage();
        if (!p) return { success: false, error: 'Browser not open. Open Gmail first.' };

        try {
            // Switch to correct Gmail account slot if specified
            let gmailUrl = 'https://mail.google.com';
            if (accountEmail) {
                const idx = await this.resolveGmailIndex(accountEmail);
                if (idx >= 0) gmailUrl = `https://mail.google.com/mail/u/${idx}/`;
            }
            // Make sure Gmail is open
            const url = p.url();
            if (!url.includes('mail.google.com')) {
                await p.goto(gmailUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
                await delay(2000);
            } else if (accountEmail && !url.includes(`/u/${await this.resolveGmailIndex(accountEmail)}/`)) {
                // Wrong account slot — switch
                await p.goto(gmailUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
                await delay(2000);
            }

            // Click Compose
            const composeSelectors = [
                '[gh="cm"]',
                '.T-I.T-I-KE.L3',
                'div[class*="compose"]',
                'button:has-text("Compose")',
            ];
            let composed = false;
            for (const sel of composeSelectors) {
                try { await p.click(sel, { timeout: 3000 }); composed = true; break; } catch {}
            }
            if (!composed) return { success: false, error: 'Could not find Compose button' };
            await delay(1000);

            // To field
            await p.fill('input[name="to"]', to, { timeout: 5000 }).catch(() => {});
            await p.keyboard.press('Tab');
            await delay(300);

            // Subject field
            await p.fill('input[name="subjectbox"]', subject, { timeout: 5000 }).catch(() => {});
            await delay(300);

            // Body – click in compose area
            const bodySelectors = [
                'div[aria-label="Message Body"]',
                'div.Am.Al.editable.LW-avf',
                'div[contenteditable="true"]',
            ];
            for (const sel of bodySelectors) {
                try { await p.click(sel, { timeout: 3000 }); break; } catch {}
            }
            await p.keyboard.type(body);
            await delay(300);

            // Send button
            const sendSelectors = [
                'div[aria-label="Send ‪(Ctrl-Enter)‬"]',
                'div[data-tooltip="Send"]',
                '.T-I.J-J5-Ji.aoO.T-I-atl.L3',
            ];
            for (const sel of sendSelectors) {
                try { await p.click(sel, { timeout: 3000 }); break; } catch {}
            }

            await delay(2000);
            return { success: true, message: `Email sent to ${to}: "${subject}"` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * YouTube control (play/pause/search)
     */
    async youtubeSearch(query) {
        const p = await getOrCreatePage();
        if (!p) {
            exec(`start "" "https://www.youtube.com/results?search_query=${encodeURIComponent(query)}"`);
            return { success: true, message: `Opened YouTube search for: ${query}` };
        }
        try {
            await p.goto(`https://www.youtube.com/results?search_query=${encodeURIComponent(query)}`, { waitUntil: 'domcontentloaded' });
            const results = await p.evaluate(() => {
                const vids = [];
                document.querySelectorAll('ytd-video-renderer a#video-title').forEach(a => {
                    if (vids.length < 5) vids.push({ title: a.innerText, url: 'https://youtube.com' + a.getAttribute('href') });
                });
                return vids;
            });
            return { success: true, results, message: `Found ${results.length} YouTube results for: ${query}` };
        } catch (e) {
            return { success: false, error: e.message };
        }
    }

    /**
     * Check if Playwright is available.
     */
    isAvailable() { return playwrightAvailable; }

    /**
     * Install Playwright and browsers (runs npm install).
     */
    async installPlaywright() {
        return new Promise((resolve) => {
            const appDir = path.join(__dirname, '..', '..');
            const child = spawn('cmd.exe', ['/c', 'npm install playwright && npx playwright install chromium'], {
                cwd: appDir,
                stdio: 'pipe',
                shell: true,
            });
            let output = '';
            child.stdout.on('data', d => { output += d.toString(); });
            child.stderr.on('data', d => { output += d.toString(); });
            child.on('close', (code) => {
                if (code === 0) {
                    // Reload playwright
                    try {
                        playwright = require('playwright');
                        chromium   = playwright.chromium;
                        playwrightAvailable = true;
                    } catch {}
                    resolve({ success: true, message: 'Playwright installed. Browser automation now available.' });
                } else {
                    resolve({ success: false, error: output.slice(-500), message: 'Playwright install failed' });
                }
            });
        });
    }
}

module.exports = new BrowserAutomation();

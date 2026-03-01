// =============================================================
// Pecifics AI — Renderer (Multi-Task, Vision, User Input)
// =============================================================

(function () {
    'use strict';

    // ─── DOM ────────────────────────────────────────────────────
    const $ = (s) => document.querySelector(s);
    const $$ = (s) => document.querySelectorAll(s);

    const chatContainer   = $('#chatContainer');
    const messageInput    = $('#messageInput');
    const sendBtn         = $('#sendBtn');
    const stopBtn         = $('#stopBtn');
    const progressArea    = $('#progressArea');
    const progressFill    = $('#progressFill');
    const progressText    = $('#progressText');
    const statusDot       = $('#statusDot');
    const previewToggle   = $('#previewToggle');
    const modeLabel       = $('#modeLabel');
    const welcomeMessage  = $('#welcomeMessage');
    const offlineBanner   = $('#offlineBanner');
    const retryBtn        = $('#retryBtn');

    // Settings
    const settingsBtn     = $('#settingsBtn');
    const backBtn         = $('#backBtn');
    const chatView        = $('#chatView');
    const settingsView    = $('#settingsView');
    const colabUrl        = $('#colabUrl');
    const cogagentUrl     = $('#cogagentUrl');
    const testConnectionBtn = $('#testConnectionBtn');
    const connectionStatus  = $('#connectionStatus');
    const saveSettingsBtn   = $('#saveSettingsBtn');
    const minimizeBtn     = $('#minimizeBtn');
    const closeBtn        = $('#closeBtn');
    const qualityRange    = $('#screenshotQuality');
    const qualityValue    = $('#qualityValue');

    // ─── STATE ──────────────────────────────────────────────────
    let backendUrl = 'http://localhost:8000';
    let conversationHistory = [];
    let isProcessing = false;
    let shouldStop = false;
    let previewMode = false;
    let screenWidth = 1920;
    let screenHeight = 1080;
    let userHome = '';

    // ─── INIT ───────────────────────────────────────────────────
    async function init() {
        try {
            const settings = await window.electronAPI.getSettings();
            if (settings.colabUrl) backendUrl = settings.colabUrl;
            if (settings.colabUrl) colabUrl.value = settings.colabUrl;
            if (settings.cogagentUrl) cogagentUrl.value = settings.cogagentUrl;
            if (settings.alwaysOnTop !== undefined) {
                $('#alwaysOnTop').checked = settings.alwaysOnTop;
            }

            // Push saved CogAgent URL to backend on startup so it persists in memory
            if (settings.cogagentUrl) {
                try {
                    await fetch(`${backendUrl}/set_config`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ cogagent_url: settings.cogagentUrl }),
                    });
                    console.log('[init] Pushed CogAgent URL to backend:', settings.cogagentUrl);
                } catch (e) { console.warn('[init] Could not push CogAgent URL to backend:', e); }

                // Also verify direct CogAgent connectivity
                try {
                    const cogHealth = await window.electronAPI.cogagentHealth();
                    if (cogHealth.ok) {
                        console.log('[init] CogAgent DIRECT connection OK:', cogHealth.data);
                    } else {
                        console.warn('[init] CogAgent not reachable directly:', cogHealth.error);
                    }
                } catch (e) { console.warn('[init] CogAgent health check failed:', e); }
            }
        } catch (e) {}

        try {
            const info = await window.electronAPI.getScreenInfo();
            if (info) { screenWidth = info.width; screenHeight = info.height; }
        } catch (e) {}

        try { userHome = await window.electronAPI.getUserHome(); } catch (e) {}

        checkBackendConnection();
        setInterval(checkBackendConnection, 30000);
    }

    // ─── BACKEND CONNECTION ─────────────────────────────────────
    async function checkBackendConnection() {
        let langchainOk = false;
        try {
            const resp = await fetch(`${backendUrl}/health`, { signal: AbortSignal.timeout(5000) });
            if (resp.ok) langchainOk = true;
        } catch (e) {}

        // Also check CogAgent direct reachability
        let cogOk = false;
        const currentCogUrl = cogagentUrl.value.trim();
        if (currentCogUrl) {
            try {
                const cogHealth = await window.electronAPI.cogagentHealth();
                cogOk = cogHealth.ok;
            } catch (e) {}
        }

        // Update status indicator
        if (langchainOk || cogOk) {
            statusDot.classList.add('connected');
            const parts = [];
            if (langchainOk) parts.push('LLM Backend');
            if (cogOk) parts.push('CogAgent');
            statusDot.title = `Connected: ${parts.join(' + ')}`;
            offlineBanner.classList.remove('visible');
            return true;
        }

        statusDot.classList.remove('connected');
        statusDot.title = 'Disconnected — start backend and/or set CogAgent URL';
        offlineBanner.classList.add('visible');
        return false;
    }

    retryBtn.addEventListener('click', () => checkBackendConnection());

    // ─── MESSAGES ───────────────────────────────────────────────
    function addMessage(text, role = 'ai') {
        if (welcomeMessage) welcomeMessage.style.display = 'none';

        const msg = document.createElement('div');
        msg.className = `msg ${role}`;

        if (role === 'ai') {
            msg.innerHTML = `
                <div class="msg-avatar">P</div>
                <div class="msg-bubble">
                    <div class="msg-text">${escapeHtml(text)}</div>
                    <div class="msg-actions">
                        <button class="msg-action-btn copy-btn" title="Copy">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
                            Copy
                        </button>
                    </div>
                </div>`;
        } else {
            msg.innerHTML = `
                <div class="msg-bubble">
                    <div class="msg-text">${escapeHtml(text)}</div>
                    <div class="msg-actions">
                        <button class="msg-action-btn copy-btn" title="Copy">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
                            Copy
                        </button>
                        <button class="msg-action-btn edit-btn" title="Edit & resend">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.12 2.12 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            Edit
                        </button>
                    </div>
                </div>`;
        }

        chatContainer.appendChild(msg);
        scrollToBottom();

        // Wire copy/edit buttons
        const copyBtn = msg.querySelector('.copy-btn');
        if (copyBtn) {
            copyBtn.addEventListener('click', () => {
                navigator.clipboard.writeText(text);
                copyBtn.classList.add('copied');
                copyBtn.innerHTML = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> Copied`;
                setTimeout(() => {
                    copyBtn.classList.remove('copied');
                    copyBtn.innerHTML = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg> Copy`;
                }, 2000);
            });
        }

        const editBtn = msg.querySelector('.edit-btn');
        if (editBtn) {
            editBtn.addEventListener('click', () => startEdit(msg, text));
        }

        return msg;
    }

    function addThinking() {
        if (welcomeMessage) welcomeMessage.style.display = 'none';
        const msg = document.createElement('div');
        msg.className = 'msg ai';
        msg.id = 'thinkingMsg';
        msg.innerHTML = `
            <div class="msg-avatar">P</div>
            <div class="msg-bubble">
                <div class="thinking">
                    <div class="thinking-dot"></div>
                    <div class="thinking-dot"></div>
                    <div class="thinking-dot"></div>
                </div>
            </div>`;
        chatContainer.appendChild(msg);
        scrollToBottom();
        return msg;
    }

    function removeThinking() {
        const t = $('#thinkingMsg');
        if (t) t.remove();
    }

    // ─── EDIT USER MESSAGE ──────────────────────────────────────
    function startEdit(msgEl, originalText) {
        const bubble = msgEl.querySelector('.msg-bubble');
        bubble.classList.add('editing');
        bubble.innerHTML = `
            <textarea class="edit-textarea" rows="3">${escapeHtml(originalText)}</textarea>
            <div class="edit-actions">
                <button class="edit-save">Resend</button>
                <button class="edit-cancel">Cancel</button>
            </div>`;

        const textarea = bubble.querySelector('.edit-textarea');
        textarea.focus();
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);

        bubble.querySelector('.edit-save').addEventListener('click', () => {
            const newText = textarea.value.trim();
            if (newText) {
                // Remove all messages after this one
                let next = msgEl.nextElementSibling;
                while (next) {
                    const toRemove = next;
                    next = next.nextElementSibling;
                    toRemove.remove();
                }
                // Trim conversation history
                const idx = conversationHistory.findLastIndex(h => h.role === 'user' && h.content === originalText);
                if (idx !== -1) conversationHistory.splice(idx);
                msgEl.remove();
                sendMessage(newText);
            }
        });

        bubble.querySelector('.edit-cancel').addEventListener('click', () => {
            bubble.classList.remove('editing');
            bubble.innerHTML = `
                <div class="msg-text">${escapeHtml(originalText)}</div>
                <div class="msg-actions">
                    <button class="msg-action-btn copy-btn" title="Copy">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
                        Copy
                    </button>
                    <button class="msg-action-btn edit-btn" title="Edit & resend">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.12 2.12 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        Edit
                    </button>
                </div>`;
            bubble.querySelector('.copy-btn').addEventListener('click', () => {
                navigator.clipboard.writeText(originalText);
            });
            bubble.querySelector('.edit-btn').addEventListener('click', () => startEdit(msgEl, originalText));
        });
    }

    // ─── TASK CARDS ─────────────────────────────────────────────
    function renderTaskCards(tasks) {
        const container = document.createElement('div');
        container.className = 'task-container';

        tasks.forEach((task, idx) => {
            const card = document.createElement('div');
            card.className = 'task-card';
            card.id = `task-card-${task.id}`;
            card.innerHTML = `
                <div class="task-header">
                    <div class="task-number">${task.id}</div>
                    <div class="task-title">${escapeHtml(task.description)}</div>
                    <div class="task-status">${task.needs_input ? 'Needs Info' : 'Pending'}</div>
                </div>
                <div class="task-actions-list"></div>`;
            container.appendChild(card);
        });

        chatContainer.appendChild(container);
        scrollToBottom();
        return container;
    }

    function updateTaskCard(taskId, status, statusText) {
        const card = $(`#task-card-${taskId}`);
        if (!card) return;
        card.className = `task-card ${status}`;
        const st = card.querySelector('.task-status');
        if (st) st.textContent = statusText || status;
    }

    function addActionToTaskCard(taskId, actionDesc, state = 'running') {
        const card = $(`#task-card-${taskId}`);
        if (!card) return;
        const list = card.querySelector('.task-actions-list');
        if (!list) return;
        const item = document.createElement('div');
        item.className = `task-action-item ${state}`;
        item.innerHTML = `<span class="action-icon">${state === 'running' ? '<div class="spinner-sm"></div>' : state === 'complete' ? '✓' : '✗'}</span> ${escapeHtml(actionDesc)}`;
        list.appendChild(item);
        scrollToBottom();
        return item;
    }

    function completeActionItem(item, success = true) {
        if (!item) return;
        item.className = `task-action-item ${success ? 'complete' : 'failed'}`;
        const icon = item.querySelector('.action-icon');
        if (icon) icon.innerHTML = success ? '✓' : '✗';
    }

    // ─── USER INPUT FORM ────────────────────────────────────────
    function showInputForm(task) {
        return new Promise((resolve) => {
            const form = document.createElement('div');
            form.className = 'input-form-card';
            let fieldsHtml = `<h4>📝 Input needed: ${escapeHtml(task.description)}</h4>`;

            (task.input_fields || []).forEach(field => {
                const isTextarea = field.key === 'content' || field.key === 'body' || field.key === 'message';
                fieldsHtml += `
                    <div class="form-field">
                        <label>${escapeHtml(field.label || field.key)}</label>
                        ${isTextarea
                            ? `<textarea data-key="${field.key}" placeholder="${escapeHtml(field.placeholder || '')}">${escapeHtml(field.default || '')}</textarea>`
                            : `<input type="text" data-key="${field.key}" placeholder="${escapeHtml(field.placeholder || '')}" value="${escapeHtml(field.default || '')}">`
                        }
                    </div>`;
            });

            fieldsHtml += `
                <div class="form-actions">
                    <button class="form-submit">Submit & Continue</button>
                    <button class="form-skip">Skip Task</button>
                </div>`;

            form.innerHTML = fieldsHtml;
            chatContainer.appendChild(form);
            scrollToBottom();

            // Focus first input
            const firstInput = form.querySelector('input, textarea');
            if (firstInput) firstInput.focus();

            form.querySelector('.form-submit').addEventListener('click', () => {
                const values = {};
                form.querySelectorAll('[data-key]').forEach(el => {
                    values[el.dataset.key] = el.value;
                });
                form.remove();
                resolve({ submitted: true, values });
            });

            form.querySelector('.form-skip').addEventListener('click', () => {
                form.remove();
                resolve({ submitted: false, values: {} });
            });
        });
    }

    // ─── SEND MESSAGE ───────────────────────────────────────────
    async function sendMessage(text) {
        if (!text || isProcessing) return;
        isProcessing = true;
        shouldStop = false;
        sendBtn.disabled = true;

        addMessage(text, 'user');
        conversationHistory.push({ role: 'user', content: text });

        const thinkingEl = addThinking();
        showProgress('Planning tasks...', 5);

        try {
            await window.electronAPI.resetStopFlag();

            // Take screenshot
            let screenshotB64 = null;
            try {
                const ssData = await window.electronAPI.takeScreenshot();
                if (ssData && ssData.screenshot) screenshotB64 = ssData.screenshot;
            } catch (e) {
                console.warn('Screenshot failed:', e);
            }

            // Call backend
            const resp = await fetch(`${backendUrl}/chat`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    message: text,
                    screenshot: screenshotB64,
                    conversation_history: conversationHistory.slice(-10),
                    screen_width: screenWidth,
                    screen_height: screenHeight,
                    user_home: userHome,
                }),
            });

            if (!resp.ok) {
                throw new Error(`Backend error: ${resp.status} ${resp.statusText}`);
            }

            const data = await resp.json();
            removeThinking();

            // Show AI message
            if (data.message) {
                addMessage(data.message, 'ai');
                conversationHistory.push({ role: 'assistant', content: data.message });
            }

            const tasks = data.tasks || [];
            if (tasks.length === 0) {
                hideProgress();
                isProcessing = false;
                sendBtn.disabled = false;
                return;
            }

            // Render task cards
            renderTaskCards(tasks);

            // Execute tasks sequentially
            await executeTasks(tasks, data.expected_result, text);

        } catch (err) {
            removeThinking();
            addMessage(`Error: ${err.message}`, 'ai');
            console.error('Send error:', err);
        }

        hideProgress();
        isProcessing = false;
        sendBtn.disabled = false;
    }

    // ─── EXECUTE TASKS ──────────────────────────────────────────
    async function executeTasks(tasks, expectedResult, originalMessage) {
        const totalTasks = tasks.length;

        for (let i = 0; i < tasks.length; i++) {
            if (shouldStop) {
                addMessage('⏹ Execution stopped by user.', 'ai');
                break;
            }

            const task = tasks[i];
            const pct = Math.round(((i) / totalTasks) * 100);
            showProgress(`Task ${task.id}/${totalTasks}: ${task.description}`, pct);
            updateTaskCard(task.id, 'running', 'Running...');

            // Handle needs_input tasks
            if (task.needs_input && task.input_fields && task.input_fields.length > 0) {
                updateTaskCard(task.id, 'running', 'Waiting for input...');
                const inputResult = await showInputForm(task);

                if (!inputResult.submitted) {
                    updateTaskCard(task.id, 'error', 'Skipped');
                    continue;
                }

                // Re-plan this task with user's input
                const inputDesc = Object.entries(inputResult.values)
                    .map(([k, v]) => `${k}: ${v}`)
                    .join(', ');

                try {
                    const resp = await fetch(`${backendUrl}/chat`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            message: `${task.description}. User provided: ${inputDesc}`,
                            conversation_history: conversationHistory.slice(-6),
                            screen_width: screenWidth,
                            screen_height: screenHeight,
                            user_home: userHome,
                            user_choice: { type: 'input', value: inputResult.values },
                        }),
                    });
                    const reData = await resp.json();
                    // Replace task actions with new ones
                    const reTasks = reData.tasks || [];
                    if (reTasks.length > 0 && reTasks[0].actions) {
                        task.actions = reTasks[0].actions;
                        task.needs_input = false;
                    }
                } catch (e) {
                    console.error('Re-plan error:', e);
                    updateTaskCard(task.id, 'error', 'Re-plan failed');
                    continue;
                }
            }

            // Execute task actions
            let taskSuccess = true;
            const actions = task.actions || [];

            for (let j = 0; j < actions.length; j++) {
                if (shouldStop) break;

                const action = actions[j];
                const actionName = action.name || action.action || 'unknown';
                const actionParams = action.parameters || action.params || {};
                const actionDesc = actionParams.goal || actionParams.description || actionName;

                showProgress(`Task ${task.id}: ${actionDesc}`, pct + Math.round(((j + 1) / actions.length) * (100 / totalTasks)));

                if (actionName === 'vision_task') {
                    // Vision task — screenshot loop
                    const result = await executeVisionTask(
                        actionParams.goal || task.description,
                        actionParams.max_steps || 30,
                        task.id
                    );
                    if (!result.success) taskSuccess = false;

                } else if (actionName === 'generate_ppt') {
                    // PPT generation — route through action-executor so shell.openPath fires
                    const pptTopic = actionParams.topic || actionParams.title || task.description;
                    const item = addActionToTaskCard(task.id, `Creating PPT: ${pptTopic}`);
                    showProgress(`Generating presentation: ${pptTopic}...`, pct + 10);
                    try {
                        const result = await window.electronAPI.executeAction({
                            action: 'generate_ppt',
                            params: actionParams,
                        });
                        completeActionItem(item, result && result.success !== false);
                        if (result && result.success !== false) {
                            addMessage(`✅ Presentation created and opened:\n${result.path || ''}`, 'ai');
                        } else {
                            addMessage(`❌ PPT error: ${(result && result.error) || 'Unknown error'}`, 'ai');
                            taskSuccess = false;
                        }
                    } catch (e) {
                        completeActionItem(item, false);
                        addMessage(`❌ PPT error: ${e.message}`, 'ai');
                        taskSuccess = false;
                    }

                } else {
                    // Regular action — execute via Electron
                    const item = addActionToTaskCard(task.id, actionDesc);

                    if (previewMode) {
                        completeActionItem(item, true);
                        await delay(300);
                        continue;
                    }

                    try {
                        const result = await window.electronAPI.executeAction({
                            action: actionName,
                            params: actionParams,
                        });
                        const success = result && result.success !== false;
                        completeActionItem(item, success);
                        if (!success) taskSuccess = false;

                        // Small delay between actions
                        await delay(500);
                    } catch (e) {
                        completeActionItem(item, false);
                        taskSuccess = false;
                        console.error(`Action ${actionName} error:`, e);
                    }
                }
            }

            updateTaskCard(task.id, taskSuccess ? 'done' : 'error', taskSuccess ? 'Done' : 'Failed');
        }

        // Final progress
        showProgress('All tasks complete', 100);

        // Verify if we have expected result
        if (expectedResult && !shouldStop) {
            await verifyCompletion(originalMessage, expectedResult);
        }

        setTimeout(hideProgress, 2000);
    }

    // ─── VISION TASK LOOP ───────────────────────────────────────
    async function executeVisionTask(goal, maxSteps = 30, taskId) {
        const currentCogUrl = cogagentUrl.value.trim();
        const useDirect = !!currentCogUrl;          // direct CogAgent when URL is set

        if (!currentCogUrl) {
            console.warn('[vision] No CogAgent URL set! Go to Settings and paste your Kaggle ngrok URL.');
            addActionToTaskCard(taskId, '⚠️ No CogAgent URL set — open Settings and paste your Kaggle ngrok URL');
        } else {
            console.log('[vision] Using DIRECT CogAgent connection:', currentCogUrl);
        }

        const stepHistory = [];
        let stepCount = 0;

        showVisionOverlay(goal);

        while (stepCount < maxSteps) {
            if (shouldStop) {
                hideVisionOverlay();
                return { success: false, reason: 'Stopped by user' };
            }

            stepCount++;
            const item = addActionToTaskCard(taskId, `Vision step ${stepCount}: analyzing screen...`);

            try {
                // Capture hi-res screenshot
                const ssData = await window.electronAPI.takeScreenshotHires();
                if (!ssData || !ssData.screenshot) {
                    completeActionItem(item, false);
                    hideVisionOverlay();
                    return { success: false, reason: 'Screenshot failed' };
                }
                const screenshotB64 = ssData.screenshot;

                let action;

                if (useDirect) {
                    // ── DIRECT: main-process IPC → CogAgent Kaggle (120s timeout) ──
                    action = await window.electronAPI.cogagentVisionAct({
                        screenshot:    screenshotB64,
                        goal,
                        step_history:  stepHistory,
                        screen_width:  screenWidth,
                        screen_height: screenHeight,
                    });
                } else {
                    // ── FALLBACK: proxy through langchain backend ──────────────────
                    const resp = await fetch(`${backendUrl}/vision_act`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            screenshot: screenshotB64,
                            goal,
                            step_history: stepHistory,
                            screen_width: screenWidth,
                            screen_height: screenHeight,
                            cogagent_url: cogagentUrl.value.trim() || undefined,
                        }),
                    });

                    if (!resp.ok) {
                        completeActionItem(item, false);
                        continue;
                    }
                    action = await resp.json();
                }

                // Update step display
                updateVisionStep(`Step ${stepCount}: ${action.description || action.action}`);

                // Check for done/fail
                if (action.action === 'done') {
                    completeActionItem(item, true);
                    item.querySelector('.action-icon').nextSibling.textContent = ` ✅ ${action.description || 'Task complete'}`;
                    stepHistory.push({ ...action, success: true });
                    hideVisionOverlay();
                    return { success: true, steps: stepCount };
                }

                if (action.action === 'fail') {
                    completeActionItem(item, false);
                    item.querySelector('.action-icon').nextSibling.textContent = ` ❌ ${action.description || 'Cannot proceed'}`;
                    stepHistory.push({ ...action, success: false });
                    hideVisionOverlay();
                    return { success: false, reason: action.description, steps: stepCount };
                }

                // Execute the vision action via screen-agent
                const execResult = await window.electronAPI.executeAction({
                    action: 'vision_execute',
                    params: action,
                });

                const success = execResult && execResult.success !== false;
                completeActionItem(item, success);
                item.querySelector('.action-icon').nextSibling.textContent = ` ${action.description || action.action}`;

                stepHistory.push({ ...action, success });

                // Wait for screen to settle
                await delay(1200);

            } catch (e) {
                completeActionItem(item, false);
                console.error('Vision step error:', e);
                stepHistory.push({ action: 'error', description: e.message, success: false });
            }
        }

        hideVisionOverlay();
        return { success: false, reason: `Max steps (${maxSteps}) reached`, steps: stepCount };
    }

    // ─── VISION OVERLAY ─────────────────────────────────────────
    function showVisionOverlay(goal) {
        let overlay = $('.vision-overlay');
        if (!overlay) {
            overlay = document.createElement('div');
            overlay.className = 'vision-overlay';
            document.body.appendChild(overlay);
        }
        overlay.innerHTML = `
            <h4><div class="spinner-sm"></div> Vision Agent</h4>
            <div class="vision-step active">${escapeHtml(goal)}</div>`;
        overlay.classList.add('visible');
    }

    function updateVisionStep(text) {
        const overlay = $('.vision-overlay');
        if (!overlay) return;
        const step = document.createElement('div');
        step.className = 'vision-step active';
        step.textContent = text;
        // Keep only last 4 steps visible
        const steps = overlay.querySelectorAll('.vision-step');
        if (steps.length > 4) steps[0].remove();
        overlay.appendChild(step);
    }

    function hideVisionOverlay() {
        const overlay = $('.vision-overlay');
        if (overlay) overlay.classList.remove('visible');
    }

    // ─── VERIFY COMPLETION ──────────────────────────────────────
    async function verifyCompletion(task, expectedResult) {
        try {
            const ssData = await window.electronAPI.takeScreenshot();
            const screenshotB64 = ssData && ssData.screenshot ? ssData.screenshot : null;
            const resp = await fetch(`${backendUrl}/verify`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ screenshot: screenshotB64, task, expected_result: expectedResult }),
            });
            const data = await resp.json();
            if (data.success) {
                addMessage(`✅ Verified: ${data.observation || 'Task completed successfully'}`, 'ai');
            } else if (data.should_retry) {
                addMessage(`⚠️ Verification: ${data.observation}. I can retry if needed.`, 'ai');
            }
        } catch (e) {
            console.warn('Verify error:', e);
        }
    }

    // ─── PROGRESS ───────────────────────────────────────────────
    function showProgress(text, pct = 0) {
        progressArea.classList.add('visible');
        progressText.textContent = text;
        progressFill.style.width = `${Math.min(pct, 100)}%`;
    }

    function hideProgress() {
        progressArea.classList.remove('visible');
        progressFill.style.width = '0%';
        progressText.textContent = 'Ready';
    }

    // ─── SETTINGS ───────────────────────────────────────────────
    function showSettings() {
        chatView.style.display = 'none';
        chatView.classList.remove('active');
        settingsView.style.display = 'flex';
        settingsView.classList.add('active');
    }

    function showChat() {
        settingsView.style.display = 'none';
        settingsView.classList.remove('active');
        chatView.style.display = 'flex';
        chatView.classList.add('active');
    }

    settingsBtn.addEventListener('click', () => showSettings());
    backBtn.addEventListener('click', () => showChat());

    testConnectionBtn.addEventListener('click', async () => {
        const url = colabUrl.value.trim() || 'http://localhost:8000';
        const cogUrl = cogagentUrl.value.trim();
        connectionStatus.textContent = 'Testing...';
        connectionStatus.className = 'conn-status';
        connectionStatus.style.display = 'block';

        const results = [];

        // Test LangChain backend
        try {
            const resp = await fetch(`${url}/health`, { signal: AbortSignal.timeout(5000) });
            if (resp.ok) {
                const data = await resp.json();
                results.push(`✓ LLM Backend — ${data.llm}`);
            } else {
                results.push(`✗ LLM Backend — HTTP ${resp.status}`);
            }
        } catch (e) {
            results.push(`✗ LLM Backend — cannot reach ${url}`);
        }

        // Test CogAgent direct
        if (cogUrl) {
            try {
                const cogHealth = await window.electronAPI.cogagentHealth();
                if (cogHealth.ok) {
                    results.push(`✓ CogAgent — ${cogHealth.data.model || 'connected'}`);
                } else {
                    results.push(`✗ CogAgent — ${cogHealth.error}`);
                }
            } catch (e) {
                results.push(`✗ CogAgent — ${e.message}`);
            }
        } else {
            results.push('⚠ CogAgent — no URL set');
        }

        const allOk = results.every(r => r.startsWith('✓'));
        connectionStatus.textContent = results.join('  |  ');
        connectionStatus.className = `conn-status ${allOk ? 'ok' : 'fail'}`;
    });

    saveSettingsBtn.addEventListener('click', async () => {
        const settings = {
            colabUrl: colabUrl.value.trim() || 'http://localhost:8000',
            cogagentUrl: cogagentUrl.value.trim(),
            screenshotInterval: parseInt($('#screenshotInterval').value) || 1000,
            screenshotQuality: parseInt(qualityRange.value) || 80,
            autoCapture: $('#autoCaptureToggle').checked,
            alwaysOnTop: $('#alwaysOnTop').checked,
        };
        backendUrl = settings.colabUrl;
        await window.electronAPI.saveSettings(settings);
        window.electronAPI.toggleAlwaysOnTop(settings.alwaysOnTop);

        // Push CogAgent URL to backend so it persists in memory
        if (settings.cogagentUrl) {
            try {
                await fetch(`${backendUrl}/set_config`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ cogagent_url: settings.cogagentUrl }),
                });
            } catch(e) { console.warn('Could not push config to backend:', e); }
        }
        checkBackendConnection();

        // Flash save button
        saveSettingsBtn.textContent = 'Saved ✓';
        setTimeout(() => { saveSettingsBtn.textContent = 'Save Settings'; }, 1500);
    });

    qualityRange.addEventListener('input', () => {
        qualityValue.textContent = `${qualityRange.value}%`;
    });

    // ─── WINDOW CONTROLS ────────────────────────────────────────
    minimizeBtn.addEventListener('click', () => window.electronAPI.minimizeWindow());
    closeBtn.addEventListener('click', () => window.electronAPI.closeWindow());

    // ─── INPUT ──────────────────────────────────────────────────
    messageInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            const text = messageInput.value.trim();
            if (text && !isProcessing) {
                messageInput.value = '';
                messageInput.style.height = 'auto';
                sendMessage(text);
            }
        }
    });

    // Auto-resize textarea
    messageInput.addEventListener('input', () => {
        messageInput.style.height = 'auto';
        messageInput.style.height = Math.min(messageInput.scrollHeight, 120) + 'px';
    });

    sendBtn.addEventListener('click', () => {
        const text = messageInput.value.trim();
        if (text && !isProcessing) {
            messageInput.value = '';
            messageInput.style.height = 'auto';
            sendMessage(text);
        }
    });

    stopBtn.addEventListener('click', async () => {
        shouldStop = true;
        try { await window.electronAPI.stopExecution(); } catch (e) {}
        addMessage('⏹ Stopping...', 'ai');
    });

    // Preview toggle
    previewToggle.addEventListener('change', () => {
        previewMode = previewToggle.checked;
        modeLabel.textContent = previewMode ? 'Preview Mode' : 'Live Mode';
    });

    // Quick action buttons
    document.querySelectorAll('.quick-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const prompt = btn.dataset.prompt;
            if (prompt && !isProcessing) {
                sendMessage(prompt);
            }
        });
    });

    // Show settings from system event
    window.electronAPI.onShowSettings(() => {
        showSettings();
    });

    // ─── HELPERS ────────────────────────────────────────────────
    function escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function scrollToBottom() {
        requestAnimationFrame(() => {
            chatContainer.scrollTop = chatContainer.scrollHeight;
        });
    }

    function delay(ms) {
        return new Promise(r => setTimeout(r, ms));
    }

    // ─── BOOT ───────────────────────────────────────────────────
    init();

})();

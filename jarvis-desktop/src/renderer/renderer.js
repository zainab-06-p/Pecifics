// ============================================
// Pecifics Desktop Assistant - Renderer Script
// ============================================

class PecificsApp {
    constructor() {
        // State
        this.settings = {};
        this.conversationHistory = [];
        this.isProcessing = false;
        this.latestScreenshot = null;
        this.previewMode = true; // START IN PREVIEW MODE FOR SAFETY!
        this.pendingActions = null; // Store actions waiting for confirmation
        
        // Feedback loop state
        this.feedbackLoopEnabled = true;
        this.currentTask = '';
        this.expectedResult = '';
        this.maxRetries = 3;
        this.currentRetry = 0;
        
        // Interactive choice state
        this.waitingForChoice = false;
        this.choiceType = null;
        this.choiceOptions = [];
        this.pendingTask = null;
        
        // DOM Elements
        this.elements = {
            // Views
            chatView: document.getElementById('chatView'),
            settingsView: document.getElementById('settingsView'),
            
            // Chat
            chatContainer: document.getElementById('chatContainer'),
            messageInput: document.getElementById('messageInput'),
            sendBtn: document.getElementById('sendBtn'),
            welcomeMessage: document.getElementById('welcomeMessage'),
            
            // Screenshot
            screenshotImg: document.getElementById('screenshotImg'),
            noScreenshot: document.getElementById('noScreenshot'),
            autoCaptureToggle: document.getElementById('autoCaptureToggle'),
            
            // Status
            statusIndicator: document.getElementById('statusIndicator'),
            actionStatus: document.getElementById('actionStatus'),
            progressBar: document.getElementById('progressBar'),
            actionText: document.getElementById('actionText'),
            stopBtn: document.getElementById('stopBtn'),
            
            // Window controls
            settingsBtn: document.getElementById('settingsBtn'),
            minimizeBtn: document.getElementById('minimizeBtn'),
            closeBtn: document.getElementById('closeBtn'),
            
            // Settings
            colabUrl: document.getElementById('colabUrl'),
            screenshotInterval: document.getElementById('screenshotInterval'),
            screenshotQuality: document.getElementById('screenshotQuality'),
            qualityValue: document.getElementById('qualityValue'),
            hotkey: document.getElementById('hotkey'),
            alwaysOnTop: document.getElementById('alwaysOnTop'),
            testConnectionBtn: document.getElementById('testConnectionBtn'),
            connectionStatus: document.getElementById('connectionStatus'),
            backBtn: document.getElementById('backBtn'),
            saveSettingsBtn: document.getElementById('saveSettingsBtn'),
            
            // Preview Mode
            previewModeToggle: document.getElementById('previewModeToggle'),
            modeLabel: document.getElementById('modeLabel'),
            modeHint: document.getElementById('modeHint')
        };
        
        this.init();
    }
    
    async init() {
        // Load settings
        await this.loadSettings();
        
        // Setup event listeners
        this.setupEventListeners();        
        // Setup stop button
        this.setupStopButton();        
        // Setup screenshot listener
        this.setupScreenshotListener();
        
        // Check connection status
        this.checkConnection();
        
        // Initialize preview mode (ON by default for safety)
        this.setPreviewMode(true);
        
        console.log('Pecifics initialized in PREVIEW MODE (safe mode)');
    }
    
    async loadSettings() {
        this.settings = await window.electronAPI.getSettings();
        this.applySettingsToUI();
    }
    
    applySettingsToUI() {
        this.elements.colabUrl.value = this.settings.colabUrl || '';
        this.elements.screenshotInterval.value = this.settings.screenshotInterval || 1000;
        this.elements.screenshotQuality.value = this.settings.screenshotQuality || 80;
        this.elements.qualityValue.textContent = `${this.settings.screenshotQuality || 80}%`;
        this.elements.hotkey.value = this.settings.hotkey || 'CommandOrControl+Shift+J';
        this.elements.autoCaptureToggle.checked = this.settings.autoScreenshot !== false;
    }
    
    setupEventListeners() {
        // Send message
        this.elements.sendBtn.addEventListener('click', () => this.sendMessage());
        this.elements.messageInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                this.sendMessage();
            }
        });
        
        // Auto-resize textarea
        this.elements.messageInput.addEventListener('input', () => {
            this.elements.messageInput.style.height = 'auto';
            this.elements.messageInput.style.height = Math.min(this.elements.messageInput.scrollHeight, 100) + 'px';
        });
        
        // Window controls
        this.elements.minimizeBtn.addEventListener('click', () => window.electronAPI.minimizeWindow());
        this.elements.closeBtn.addEventListener('click', () => window.electronAPI.closeWindow());
        
        // Settings
        this.elements.settingsBtn.addEventListener('click', () => this.showSettings());
        this.elements.backBtn.addEventListener('click', () => this.hideSettings());
        this.elements.saveSettingsBtn.addEventListener('click', () => this.saveSettings());
        this.elements.testConnectionBtn.addEventListener('click', () => this.testConnection());
        
        // Screenshot toggle
        this.elements.autoCaptureToggle.addEventListener('change', (e) => {
            if (e.target.checked) {
                window.electronAPI.startCapture();
            } else {
                window.electronAPI.stopCapture();
            }
        });
        
        // Quality slider
        this.elements.screenshotQuality.addEventListener('input', (e) => {
            this.elements.qualityValue.textContent = `${e.target.value}%`;
        });
        
        // Always on top
        this.elements.alwaysOnTop.addEventListener('change', (e) => {
            window.electronAPI.toggleAlwaysOnTop(e.target.checked);
        });
        
        // Preview mode toggle
        this.elements.previewModeToggle.addEventListener('change', (e) => {
            this.setPreviewMode(e.target.checked);
        });
        
        // Show settings from tray
        window.electronAPI.onShowSettings(() => this.showSettings());
    }
    
    setupStopButton() {
        this.elements.stopBtn.addEventListener('click', async () => {
            await this.stopExecution();
        });
    }
    
    showStopButton() {
        this.elements.stopBtn.style.display = 'flex';
    }
    
    hideStopButton() {
        this.elements.stopBtn.style.display = 'none';
    }
    
    async stopExecution() {
        try {
            await window.electronAPI.stopExecution();
            this.addMessage('⏹ Stopping execution...', 'system');
        } catch (error) {
            console.error('Failed to stop execution:', error);
        }
    }
    
    setupScreenshotListener() {
        window.electronAPI.onScreenshotCaptured((data) => {
            if (data && data.screenshot) {
                this.latestScreenshot = data;
                this.elements.screenshotImg.src = `data:image/jpeg;base64,${data.screenshot}`;
                this.elements.screenshotImg.classList.add('visible');
                this.elements.noScreenshot.classList.add('hidden');
            }
        });
    }
    
    // ============================================
    // Chat Functions
    // ============================================
    
    async sendMessage() {
        const message = this.elements.messageInput.value.trim();
        if (!message || this.isProcessing) return;
        
        // Check if connected
        if (!this.settings.colabUrl) {
            this.addMessage('Please configure the Colab backend URL in settings first.', 'system');
            this.showSettings();
            return;
        }
        
        // Hide welcome message
        this.elements.welcomeMessage.classList.add('hidden');
        
        // Clear input
        this.elements.messageInput.value = '';
        this.elements.messageInput.style.height = 'auto';
        
        // Add user message
        this.addMessage(message, 'user');
        this.conversationHistory.push({ role: 'user', content: message });
        
        // Show typing indicator
        this.isProcessing = true;
        this.showTypingIndicator();
        this.showActionStatus('Processing your request...');
        
        try {
            // Get fresh screenshot if auto-capture is on
            let screenshot = null;
            if (this.elements.autoCaptureToggle.checked) {
                const screenshotData = await window.electronAPI.takeScreenshot();
                if (screenshotData) {
                    screenshot = screenshotData.screenshot;
                    this.latestScreenshot = screenshotData;
                }
            } else if (this.latestScreenshot) {
                screenshot = this.latestScreenshot.screenshot;
            }
            
            // Get screen info
            const screenInfo = await window.electronAPI.getScreenInfo();
            
            // Send to backend
            const response = await this.callBackend('/chat', {
                message: message,
                screenshot: screenshot,
                conversation_history: this.conversationHistory.slice(-10),
                screen_width: screenInfo.width,
                screen_height: screenInfo.height
            });
            
            // Hide typing indicator
            this.hideTypingIndicator();
            
            if (response.error) {
                this.addMessage(`Error: ${response.error}`, 'error');
            } else {
                // CHECK FOR CLARIFICATION NEEDED
                if (response.clarification_needed) {
                    await this.handleClarificationRequest(response);
                    return;
                }

                // CHECK FOR INTERACTIVE CHOICE
                if (response.requires_choice) {
                    // AI needs user to make a choice
                    await this.handleChoiceRequest(response);
                    return; // Wait for user to choose
                }
                
                // Add assistant response
                if (response.message) {
                    // Show task count if multiple tasks detected
                    const taskInfo = response.task_count > 1 
                        ? ` (${response.task_count} tasks detected)` 
                        : '';
                    this.addMessage(response.message + taskInfo, 'assistant');
                }
                
                // Store task info for feedback loop
                this.currentTask = message;
                this.expectedResult = response.expected_result || '';
                this.currentRetry = 0;
                
                // Store task summary for display after execution
                this.currentTaskSummary = response.task_summary || null;

                // Handle actions based on mode
                if (response.actions && response.actions.length > 0) {
                    if (this.previewMode) {
                        // PREVIEW MODE: Show actions without executing
                        this.showPreviewActions(response.actions, response.task_count);
                    } else {
                        // LIVE MODE: Execute actions
                        await this.executeActions(response.actions);
                    }
                }
                
                this.conversationHistory.push({ 
                    role: 'assistant', 
                    content: response.message || 'Actions executed.'
                });
            }
            
        } catch (error) {
            this.hideTypingIndicator();
            this.addMessage(`Connection error: ${error.message}`, 'error');
        } finally {
            this.isProcessing = false;
            this.hideActionStatus();
        }
    }
    
    renderMarkdown(text) {
        if (!text) return '';
        // Bold **text**
        text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
        // Italic *text*
        text = text.replace(/\*(.+?)\*/g, '<em>$1</em>');
        // Inline code `code`
        text = text.replace(/`([^`]+)`/g, '<code>$1</code>');
        // Newlines
        text = text.replace(/\n/g, '<br>');
        // Numbered list items: 1. item
        text = text.replace(/^(\d+\.\s)/gm, '<span style="opacity:0.7">$1</span>');
        // Bullet list items
        text = text.replace(/^[-•]\s/gm, '• ');
        return text;
    }

    addMessage(content, type, actions = null) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}`;
        if (type === 'assistant' || type === 'system') {
            messageDiv.innerHTML = this.renderMarkdown(content);
        } else {
            messageDiv.textContent = content;
        }
        
        // Add action cards if present
        if (actions && actions.length > 0) {
            actions.forEach(action => {
                const actionCard = document.createElement('div');
                actionCard.className = 'action-card';
                actionCard.innerHTML = `
                    <div class="action-name">🔧 ${action.name || action.function?.name || 'Action'}</div>
                    <div class="action-params">${JSON.stringify(action.parameters || action.arguments || {}, null, 2)}</div>
                `;
                messageDiv.appendChild(actionCard);
            });
        }
        
        this.elements.chatContainer.appendChild(messageDiv);
        this.scrollToBottom();
    }
    
    showTypingIndicator() {
        const indicator = document.createElement('div');
        indicator.className = 'typing-indicator';
        indicator.id = 'typingIndicator';
        indicator.innerHTML = '<span></span><span></span><span></span>';
        this.elements.chatContainer.appendChild(indicator);
        this.scrollToBottom();
    }
    
    hideTypingIndicator() {
        const indicator = document.getElementById('typingIndicator');
        if (indicator) indicator.remove();
    }
    
    showActionStatus(text) {
        this.elements.actionStatus.classList.add('active');
        this.elements.progressBar.classList.add('indeterminate');
        this.elements.actionText.textContent = text;
    }
    
    hideActionStatus() {
        this.elements.actionStatus.classList.remove('active');
        this.elements.progressBar.classList.remove('indeterminate');
    }
    
    scrollToBottom() {
        this.elements.chatContainer.scrollTop = this.elements.chatContainer.scrollHeight;
    }
    
    // ============================================
    // Preview Mode Functions
    // ============================================
    
    setPreviewMode(enabled) {
        this.previewMode = enabled;
        this.elements.previewModeToggle.checked = enabled;
        
        if (enabled) {
            this.elements.modeLabel.textContent = '👁️ Preview Mode';
            this.elements.modeLabel.classList.add('preview');
            this.elements.modeHint.textContent = 'Actions shown but NOT executed';
        } else {
            this.elements.modeLabel.textContent = '🟢 Live Mode';
            this.elements.modeLabel.classList.remove('preview');
            this.elements.modeHint.textContent = 'Actions WILL be executed';
        }
    }
    
    showPreviewActions(actions, taskCount = 1) {
        // Store pending actions
        this.pendingActions = actions;
        
        // Create preview container
        const previewDiv = document.createElement('div');
        previewDiv.className = 'preview-actions-container';
        previewDiv.id = 'previewActionsContainer';
        
        const taskInfo = taskCount > 1 ? ` (${taskCount} tasks)` : '';
        
        let html = `
            <div class="preview-banner">👁️ PREVIEW MODE - Actions will NOT be executed</div>
            <div class="preview-actions-header">Pecifics wants to perform ${actions.length} action(s)${taskInfo}:</div>
        `;
        
        actions.forEach((action, index) => {
            const actionName = action.name || action.function?.name || 'Unknown';
            const params = action.parameters || action.arguments || {};
            
            // Check if action would be safe
            const safetyIcon = this.getSafetyIcon(actionName, params);
            
            html += `
                <div class="preview-action-item">
                    <span class="preview-action-number">${index + 1}.</span>
                    <span class="preview-action-name">${safetyIcon} ${actionName}</span>
                    <div class="preview-action-params">${JSON.stringify(params, null, 2)}</div>
                </div>
            `;
        });
        
        html += `
            <button class="execute-preview-btn" id="executePreviewBtn">
                ✅ Approve & Execute These Actions
            </button>
            <button class="cancel-preview-btn" id="cancelPreviewBtn">
                ❌ Cancel - Don't Execute
            </button>
        `;
        
        previewDiv.innerHTML = html;
        this.elements.chatContainer.appendChild(previewDiv);
        this.scrollToBottom();
        
        // Add event listeners
        document.getElementById('executePreviewBtn').addEventListener('click', () => {
            this.executePreviewedActions();
        });
        
        document.getElementById('cancelPreviewBtn').addEventListener('click', () => {
            this.cancelPreviewedActions();
        });
    }
    
    getSafetyIcon(actionName, params) {
        // Visual indicator of action safety
        const dangerousActions = ['delete_file', 'delete_folder', 'run_command'];
        const cautionActions = ['move_file', 'rename_file', 'click_at', 'type_text'];
        
        if (dangerousActions.includes(actionName)) {
            return '🔴';
        } else if (cautionActions.includes(actionName)) {
            return '🟡';
        } else {
            return '🟢';
        }
    }
    
    async executePreviewedActions() {
        if (!this.pendingActions) return;
        
        // Remove preview container
        const previewContainer = document.getElementById('previewActionsContainer');
        if (previewContainer) {
            previewContainer.remove();
        }
        
        // Add confirmation message
        this.addMessage('✅ Actions approved! Executing...', 'system');
        
        // Execute the actions
        await this.executeActions(this.pendingActions);
        
        // Clear pending actions
        this.pendingActions = null;
    }
    
    cancelPreviewedActions() {
        // Remove preview container
        const previewContainer = document.getElementById('previewActionsContainer');
        if (previewContainer) {
            previewContainer.remove();
        }
        
        // Add cancellation message
        this.addMessage('❌ Actions cancelled. Nothing was executed.', 'system');
        
        // Clear pending actions
        this.pendingActions = null;
    }

    // ============================================
    // Clarification System
    // ============================================

    async handleClarificationRequest(response) {
        this.hideTypingIndicator();
        // Show the AI's question message
        if (response.message) {
            this.addMessage(response.message, 'assistant');
        }
        // Store the pending response context
        this.pendingClarificationResponse = response;
        // Show the input form
        this.showClarificationForm(
            response.clarification_question || 'Please provide the following details:',
            response.clarification_fields || []
        );
    }

    showClarificationForm(question, fields) {
        const formDiv = document.createElement('div');
        formDiv.className = 'clarification-form';
        formDiv.id = 'clarificationForm';

        let fieldsHtml = '';
        fields.forEach(f => {
            fieldsHtml += `
                <div class="clarification-field">
                    <label>${f.label || f.key}</label>
                    <input type="text"
                           id="clarify_${f.key}"
                           placeholder="${f.placeholder || ''}"
                           value="${f.default || ''}" />
                </div>`;
        });

        formDiv.innerHTML = `
            <div class="clarification-question">${question}</div>
            <div class="clarification-fields">${fieldsHtml}</div>
            <div class="clarification-actions">
                <button id="clarifySubmitBtn" class="clarification-submit-btn">Submit ↵</button>
                <button id="clarifyCancelBtn" class="clarification-cancel-btn">Cancel</button>
            </div>`;

        this.elements.chatContainer.appendChild(formDiv);
        this.scrollToBottom();

        // Focus first input
        const firstInput = formDiv.querySelector('input');
        if (firstInput) setTimeout(() => firstInput.focus(), 100);

        const fields_ = fields;
        document.getElementById('clarifySubmitBtn').addEventListener('click', () => {
            this.submitClarificationForm(fields_);
        });
        document.getElementById('clarifyCancelBtn').addEventListener('click', () => {
            formDiv.remove();
            this.addMessage('❌ Clarification cancelled.', 'system');
            this.pendingClarificationResponse = null;
        });
        // Allow Enter key on last input
        formDiv.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') this.submitClarificationForm(fields_);
        });
    }

    async submitClarificationForm(fields) {
        const formDiv = document.getElementById('clarificationForm');

        // Collect values
        const answers = {};
        fields.forEach(f => {
            const input = document.getElementById(`clarify_${f.key}`);
            answers[f.key] = input ? input.value.trim() || f.default || '' : f.default || '';
        });

        // Remove the form
        if (formDiv) formDiv.remove();

        // ── Special: 2FA / OTP code ── type directly into the browser page ──
        if (answers.otp_code) {
            this.addMessage(`🔐 Entering 2FA code into browser...`, 'system');
            await window.electronAPI.executeAction({
                action: 'browser_search_in_page', params: { text: answers.otp_code }
            });
            await this.delay(2000);
            await this.watchBrowserAfterAction();
            return;
        }

        // Build a natural-language answer message
        const answerParts = fields.map(f => `${f.label || f.key}: ${answers[f.key]}`).join(', ');
        this.addMessage(answerParts, 'user');
        this.conversationHistory.push({ role: 'user', content: answerParts });

        // Re-send to backend with clarification answers appended to original message
        this.isProcessing = true;
        this.showTypingIndicator();
        this.showActionStatus('Processing with your details...');

        try {
            const screenInfo = await window.electronAPI.getScreenInfo();
            const response = await this.callBackend('/chat', {
                message: answerParts,
                conversation_history: this.conversationHistory.slice(-12),
                screen_width: screenInfo.width,
                screen_height: screenInfo.height
            });

            this.hideTypingIndicator();

            if (response.error) {
                this.addMessage(`Error: ${response.error}`, 'error');
            } else {
                if (response.message) this.addMessage(response.message, 'assistant');
                this.currentTaskSummary = response.task_summary || null;
                if (response.actions && response.actions.length > 0) {
                    if (this.previewMode) {
                        this.showPreviewActions(response.actions, response.task_count);
                    } else {
                        await this.executeActions(response.actions);
                    }
                }
                this.conversationHistory.push({ role: 'assistant', content: response.message || 'Done.' });
            }
        } catch (error) {
            this.hideTypingIndicator();
            this.addMessage(`Connection error: ${error.message}`, 'error');
        } finally {
            this.isProcessing = false;
            this.hideActionStatus();
            this.pendingClarificationResponse = null;
        }
    }

    showTaskSummary(summary, results) {
        const card = document.createElement('div');
        card.className = 'task-summary-card';

        const successCount = results.filter(r => r.result && r.result.success !== false).length;
        const failCount = results.length - successCount;
        const statusIcon = failCount === 0 ? '✅' : '⚠️';

        // Format summary lines
        const summaryHtml = summary
            .split('\n')
            .filter(l => l.trim())
            .map(line => `<div class="summary-line">${this.renderMarkdown(line)}</div>`)
            .join('');

        card.innerHTML = `
            <div class="task-summary-header">${statusIcon} Task Summary</div>
            <div class="task-summary-body">${summaryHtml}</div>
            <div class="task-summary-stats">${successCount} completed · ${failCount} failed</div>`;

        this.elements.chatContainer.appendChild(card);
        this.scrollToBottom();
    }

    // ============================================
    // Browser Screen Watch (post-action blocker detection)
    // ============================================

    /**
     * After a browser action, take a page screenshot via Playwright,
     * send to /check_browser_state, and auto-handle or inform user.
     * Loops up to maxPasses times to chain-handle multiple blockers.
     */
    async watchBrowserAfterAction(maxPasses = 4) {
        for (let pass = 0; pass < maxPasses; pass++) {
            // 1. Take browser page screenshot (base64) via main process
            let b64 = null;
            try {
                const r = await window.electronAPI.executeAction({
                    action: 'browser_page_screenshot_b64', params: {}
                });
                b64 = r && r.success ? r.data : null;
            } catch { break; }

            if (!b64) break;

            // 2. Ask vision backend what's on screen
            let state;
            try {
                state = await this.callBackend('/check_browser_state', { screenshot: b64 });
            } catch { break; }

            if (!state || state.state === 'clear' || state.state === 'loading') break;

            // 3. Act on the result
            if (state.can_auto_handle && state.auto_selector) {
                this.addMessage(`👁️ Detected: **${state.description}** — auto-handling...`, 'system');
                await window.electronAPI.executeAction({
                    action: 'browser_click', params: { selector: state.auto_selector }
                });
                await this.delay(1500);
                // Also try DOM-based fast handler
                await window.electronAPI.executeAction({
                    action: 'browser_auto_handle_blockers', params: {}
                });
                await this.delay(800);
                // Loop again to check if another blocker appeared
                continue;
            }

            if (state.needs_user) {
                this.addMessage(`⚠️ **Action required:** ${state.user_message || state.description}`, 'assistant');
                // If it's 2FA, show an input form for the code
                if (state.state === '2fa') {
                    this.showClarificationForm(
                        state.user_message || 'Enter the 2FA code:',
                        [{ key: 'otp_code', label: '2FA / OTP Code', placeholder: 'e.g. 123456', default: '' }]
                    );
                }
                break; // Stop watching — waiting for user
            }

            break;
        }
    }

    // ============================================
    // Action Execution
    // ============================================

    async executeActions(actions) {
        const results = [];
        
        // Reset stop flag and show stop button
        await window.electronAPI.resetStopFlag();
        this.showStopButton();
        
        for (let i = 0; i < actions.length; i++) {
            // Check if user requested stop
            const isStopped = await window.electronAPI.checkStopFlag();
            if (isStopped) {
                this.addMessage('⏹ Execution stopped by user', 'system');
                this.hideStopButton();
                return results;
            }
            
            const action = actions[i];
            const actionName = action.name || action.function?.name;
            const params = action.parameters || action.arguments || {};
            
            this.showActionStatus(`Executing: ${actionName} (${i + 1}/${actions.length})`);
            
            try {
                // Execute action
                const result = await this.executeAction(actionName, params);
                
                // Show result
                if (result && result.blocked) {
                    this.addActionResult(actionName, params, 'blocked', result.error);
                } else if (result && result.success === false) {
                    this.addActionResult(actionName, params, 'error', result.error);
                } else {
                    this.addActionResult(actionName, params, 'success', result?.message);
                }
                
                results.push({ action: actionName, result });

                // After any browser action, watch screen for blockers
                const isBrowserAction = actionName.startsWith('browser_') || actionName === 'web_login';
                if (isBrowserAction && result && result.success !== false) {
                    await this.delay(1800); // let page settle
                    await this.watchBrowserAfterAction();
                }

                // Smart delay between actions based on action type
                const delayTime = this.getActionDelay(actionName);
                if (i < actions.length - 1) {
                    await this.delay(delayTime);
                }
                
            } catch (error) {
                console.error(`Action failed: ${actionName}`, error);
                this.addActionResult(actionName, params, 'error', error.message);
                results.push({ action: actionName, error: error.message });
            }
        }
        
        // Hide stop button when done
        this.hideStopButton();
        
        // Summary
        const successful = results.filter(r => r.result && r.result.success !== false).length;
        const failed = results.length - successful;
        
        if (failed === 0) {
            this.addMessage(`✅ All ${results.length} action(s) completed successfully!`, 'system');
        } else {
            this.addMessage(`⚠️ Completed: ${successful} success, ${failed} failed`, 'system');
        }

        // Show task summary if AI provided one
        if (this.currentTaskSummary) {
            this.showTaskSummary(this.currentTaskSummary, results);
            this.currentTaskSummary = null;
        }
        
        // FEEDBACK LOOP: Verify task completion if enabled
        if (this.feedbackLoopEnabled && this.expectedResult) {
            await this.verifyAndRetry();
        }
        
        return results;
    }
    
    // ============================================
    // Feedback Loop - Verify & Retry
    // ============================================
    
    async verifyAndRetry() {
        if (!this.feedbackLoopEnabled || this.currentRetry >= this.maxRetries) {
            return;
        }
        
        this.addMessage('🔍 Verifying task completion...', 'system');
        this.showActionStatus('Verifying task...');
        
        // Take a new screenshot to verify
        await this.delay(1500); // Wait for UI to settle
        const screenshotData = await window.electronAPI.takeScreenshot();
        
        if (!screenshotData || !screenshotData.screenshot) {
            this.addMessage('⚠️ Could not capture verification screenshot', 'system');
            return;
        }
        
        try {
            // Call verify endpoint
            const verifyResponse = await this.callBackend('/verify', {
                screenshot: screenshotData.screenshot,
                task: this.currentTask,
                expected_result: this.expectedResult
            });
            
            if (verifyResponse.error) {
                this.addMessage(`⚠️ Verification error: ${verifyResponse.error}`, 'system');
                return;
            }
            
            if (verifyResponse.success) {
                this.addMessage('✅ Task verified as complete!', 'system');
                this.currentRetry = 0; // Reset
                return;
            }
            
            // Task not complete
            if (verifyResponse.should_retry && verifyResponse.retry_actions?.length > 0) {
                this.currentRetry++;
                this.addMessage(
                    `🔄 Task incomplete. ${verifyResponse.observation || ''}\nRetrying... (${this.currentRetry}/${this.maxRetries})`,
                    'system'
                );
                
                // Execute retry actions
                if (this.previewMode) {
                    this.showPreviewActions(verifyResponse.retry_actions, 1);
                } else {
                    await this.executeActions(verifyResponse.retry_actions);
                }
            } else {
                this.addMessage(
                    `⚠️ Task may be incomplete: ${verifyResponse.observation || 'Unknown state'}`,
                    'system'
                );
            }
            
        } catch (error) {
            console.error('Verification error:', error);
            this.addMessage(`⚠️ Verification failed: ${error.message}`, 'system');
        } finally {
            this.hideActionStatus();
        }
    }

    // Get appropriate delay based on action type
    getActionDelay(actionName) {
        // Actions that need more time (app opening, typing)
        const longDelayActions = [
            'open_application', 'open-app', 'open_app', 'launch_app', 'launch_application'
        ];
        
        // Actions that need medium delay (typing, clicks)
        const mediumDelayActions = [
            'type_text', 'type-text', 'type', 'type_into_app',
            'click', 'click_at', 'press_key', 'press-key', 'press', 'hotkey'
        ];
        
        if (longDelayActions.includes(actionName)) {
            return 2000; // 2 seconds for app to fully open
        } else if (mediumDelayActions.includes(actionName)) {
            return 500; // 500ms after typing/clicking
        } else {
            return 300; // Default delay
        }
    }
    
    addActionResult(actionName, params, status, message) {
        const resultDiv = document.createElement('div');
        resultDiv.className = `action-result ${status}`;
        
        let statusIcon = '✅';
        if (status === 'blocked') statusIcon = '🛡️';
        else if (status === 'error') statusIcon = '❌';
        
        resultDiv.innerHTML = `
            <div class="action-name">${statusIcon} ${actionName}</div>
            <div class="action-params">${JSON.stringify(params, null, 2)}</div>
            ${message ? `<div class="action-message">${message}</div>` : ''}
        `;
        
        this.elements.chatContainer.appendChild(resultDiv);
        this.scrollToBottom();
    }
    
    async executeAction(actionName, params) {
        // Special: get_screenshot / describe_screen — take screenshot, send to AI for vision analysis
        const visionActions = ['get_screenshot', 'describe_screen', 'what_is_on_screen', 'analyze_screen', 'screenshot_analyze'];
        if (visionActions.includes(actionName)) {
            try {
                const screenshotData = await window.electronAPI.takeScreenshot();
                if (!screenshotData || !screenshotData.screenshot) {
                    return { success: false, error: 'Could not capture screenshot' };
                }
                // Use dedicated /analyze_screen endpoint (Gemini vision)
                const response = await this.callBackend('/analyze_screen', {
                    screenshot: screenshotData.screenshot,
                    question: params.question || params.query || 'Describe exactly what is visible on my screen right now. List open windows, content, and any important details.'
                });
                if (response.error) return { success: false, error: response.error };
                const answer = response.answer || response.description || 'No description available';
                this.addMessage(`👁️ **Screen Analysis:**\n${answer}`, 'assistant');
                return { success: true, message: answer };
            } catch (e) {
                return { success: false, error: e.message };
            }
        }

        // Execute all other actions locally through main process
        try {
            const result = await window.electronAPI.executeAction({
                action: actionName,
                params: params
            });
            return result;
        } catch (error) {
            console.error(`Action ${actionName} failed:`, error);
            return {
                success: false,
                error: error.message || 'Action execution failed'
            };
        }
    }
    
    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
    
    // ============================================
    // API Communication
    // ============================================
    
    async callBackend(endpoint, data) {
        try {
            const response = await fetch(`${this.settings.colabUrl}${endpoint}`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'ngrok-skip-browser-warning': 'true'  // Bypass ngrok warning
                },
                body: JSON.stringify(data)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const contentType = response.headers.get('content-type');
            if (!contentType || !contentType.includes('application/json')) {
                throw new Error('Server returned HTML instead of JSON. Open the ngrok URL in your browser first and click "Visit Site".');
            }
            
            return await response.json();
        } catch (error) {
            return { error: error.message };
        }
    }
    
    async checkConnection() {
        if (!this.settings.colabUrl) {
            this.setConnectionStatus('disconnected');
            return;
        }
        
        this.setConnectionStatus('connecting');
        
        try {
            const response = await fetch(`${this.settings.colabUrl}/health`, {
                method: 'GET',
                headers: {
                    'ngrok-skip-browser-warning': 'true'
                },
                timeout: 5000
            });
            
            if (response.ok) {
                this.setConnectionStatus('connected');
            } else {
                this.setConnectionStatus('disconnected');
            }
        } catch (error) {
            this.setConnectionStatus('disconnected');
        }
    }
    
    setConnectionStatus(status) {
        this.elements.statusIndicator.className = 'status-indicator';
        
        switch (status) {
            case 'connected':
                this.elements.statusIndicator.classList.add('connected');
                this.elements.statusIndicator.title = 'Connected to backend';
                break;
            case 'connecting':
                this.elements.statusIndicator.classList.add('connecting');
                this.elements.statusIndicator.title = 'Connecting...';
                break;
            case 'disconnected':
                this.elements.statusIndicator.classList.add('disconnected');
                this.elements.statusIndicator.title = 'Disconnected';
                break;
        }
    }
    
    // ============================================
    // Interactive Choice Handling
    // ============================================
    
    async handleChoiceRequest(response) {
        // Execute detection actions first (e.g., detect_browsers)
        if (response.actions && response.actions.length > 0) {
            this.showActionStatus('Detecting available options...');
            
            for (const action of response.actions) {
                const actionName = action.name || action.function?.name;
                const params = action.parameters || action.arguments || {};
                
                const result = await this.executeAction(actionName, params);
                
                // Store detected options
                if (result.success) {
                    if (result.browsers) {
                        this.choiceOptions = result.browsers;
                    } else if (result.profiles) {
                        this.choiceOptions = result.profiles;
                    }
                }
            }
            
            this.hideActionStatus();
        }
        
        // Set choice state
        this.waitingForChoice = true;
        this.choiceType = response.choice_type;
        this.pendingTask = response.pending_task;
        
        // Display choice UI
        this.showChoiceUI(response.message, this.choiceOptions);
    }
    
    showChoiceUI(question, options) {
        // Create choice container
        const choiceDiv = document.createElement('div');
        choiceDiv.className = 'choice-container';
        choiceDiv.id = 'choiceContainer';
        
        let html = `
            <div class="choice-banner">🤔 AI needs your help</div>
            <div class="choice-question">${question}</div>
            <div class="choice-options">
        `;
        
        options.forEach((option, index) => {
            html += `
                <button class="choice-option-btn" data-choice="${option}">
                    ${option}
                </button>
            `;
        });
        
        html += `
            </div>
        `;
        
        choiceDiv.innerHTML = html;
        this.elements.chatContainer.appendChild(choiceDiv);
        this.scrollToBottom();
        
        // Add event listeners to all choice buttons
        document.querySelectorAll('.choice-option-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const choice = e.target.getAttribute('data-choice');
                this.handleUserChoice(choice);
            });
        });
    }
    
    async handleUserChoice(choice) {
        // Remove choice UI
        const choiceContainer = document.getElementById('choiceContainer');
        if (choiceContainer) {
            choiceContainer.remove();
        }
        
        // Add user's choice as a message
        this.addMessage(`You chose: ${choice}`, 'user');
        
        // Store the choice
        const chosenOption = choice;
        const choiceType = this.choiceType;
        const pendingTask = this.pendingTask;
        
        // Reset choice state
        this.waitingForChoice = false;
        this.choiceType = null;
        this.choiceOptions = [];
        this.pendingTask = null;
        
        // Continue the task with the chosen option
        this.isProcessing = true;
        this.showTypingIndicator();
        this.showActionStatus('Continuing with your choice...');
        
        try {
            // If we just chose a browser, now ask for profile
            if (choiceType === 'browser') {
                // Send follow-up to get profiles for chosen browser
                const followUpMessage = `I chose ${chosenOption}. ${pendingTask}`;
                
                const response = await this.callBackend('/chat', {
                    message: followUpMessage,
                    conversation_history: [
                        ...this.conversationHistory.slice(-10),
                        { role: 'user', content: `Selected browser: ${chosenOption}` }
                    ],
                    screen_width: (await window.electronAPI.getScreenInfo()).width,
                    screen_height: (await window.electronAPI.getScreenInfo()).height,
                    user_choice: { type: 'browser', value: chosenOption }
                });
                
                this.hideTypingIndicator();
                
                // Handle the response (might be another choice or action)
                if (response.requires_choice) {
                    await this.handleChoiceRequest(response);
                } else {
                    if (response.message) {
                        this.addMessage(response.message, 'assistant');
                    }
                    
                    if (response.actions && response.actions.length > 0) {
                        if (this.previewMode) {
                            this.showPreviewActions(response.actions);
                        } else {
                            await this.executeActions(response.actions);
                        }
                    }
                }
                
            } else if (choiceType === 'account') {
                // Final step - execute the task with chosen account
                const followUpMessage = `Use account: ${chosenOption}. ${pendingTask}`;
                
                const response = await this.callBackend('/chat', {
                    message: followUpMessage,
                    conversation_history: [
                        ...this.conversationHistory.slice(-10),
                        { role: 'user', content: `Selected account: ${chosenOption}` }
                    ],
                    screen_width: (await window.electronAPI.getScreenInfo()).width,
                    screen_height: (await window.electronAPI.getScreenInfo()).height,
                    user_choice: { type: 'account', value: chosenOption }
                });
                
                this.hideTypingIndicator();
                
                if (response.message) {
                    this.addMessage(response.message, 'assistant');
                }
                
                if (response.actions && response.actions.length > 0) {
                    if (this.previewMode) {
                        this.showPreviewActions(response.actions);
                    } else {
                        await this.executeActions(response.actions);
                    }
                }
            }
            
        } catch (error) {
            this.hideTypingIndicator();
            this.addMessage(`Error: ${error.message}`, 'error');
        } finally {
            this.isProcessing = false;
            this.hideActionStatus();
        }
    }
    
    // ============================================
    // Settings Functions
    // ============================================
    
    showSettings() {
        this.elements.chatView.classList.remove('active');
        this.elements.settingsView.classList.add('active');
    }
    
    hideSettings() {
        this.elements.settingsView.classList.remove('active');
        this.elements.chatView.classList.add('active');
    }
    
    async saveSettings() {
        const newSettings = {
            colabUrl: this.elements.colabUrl.value.trim().replace(/\/$/, ''), // Remove trailing slash
            screenshotInterval: parseInt(this.elements.screenshotInterval.value),
            screenshotQuality: parseInt(this.elements.screenshotQuality.value),
            autoScreenshot: this.elements.autoCaptureToggle.checked
        };
        
        await window.electronAPI.saveSettings(newSettings);
        this.settings = { ...this.settings, ...newSettings };
        
        // Check connection with new URL
        await this.checkConnection();
        
        this.hideSettings();
        this.addMessage('Settings saved!', 'system');
    }
    
    async testConnection() {
        const url = this.elements.colabUrl.value.trim().replace(/\/$/, '');
        
        if (!url) {
            this.showConnectionResult(false, 'Please enter a URL');
            return;
        }
        
        this.elements.testConnectionBtn.textContent = 'Testing...';
        this.elements.testConnectionBtn.disabled = true;
        
        try {
            const response = await fetch(`${url}/health`, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json',
                    'ngrok-skip-browser-warning': 'true'  // Skip ngrok warning page
                }
            });
            
            if (response.ok) {
                const contentType = response.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                    const data = await response.json();
                    this.showConnectionResult(true, `Connected! Status: ${data.status}`);
                } else {
                    // Got HTML instead of JSON (ngrok warning page)
                    this.showConnectionResult(false, 'Ngrok warning page detected. Open URL in browser first, click "Visit Site", then try again.');
                }
            } else {
                this.showConnectionResult(false, `Server returned: ${response.status}`);
            }
        } catch (error) {
            let errorMsg = error.message;
            if (errorMsg.includes('JSON')) {
                errorMsg = '⚠️ Ngrok free tier issue: Open the URL in your browser, click "Visit Site" button, then try again.';
            }
            this.showConnectionResult(false, errorMsg);
        } finally {
            this.elements.testConnectionBtn.textContent = 'Test Connection';
            this.elements.testConnectionBtn.disabled = false;
        }
    }
    
    showConnectionResult(success, message) {
        this.elements.connectionStatus.className = `connection-status ${success ? 'success' : 'error'}`;
        this.elements.connectionStatus.textContent = message;
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.pecifics = new PecificsApp();
});

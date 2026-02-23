// ============================================
// JARVIS Safety Guard Module
// Prevents harmful/dangerous actions
// ============================================

const path = require('path');
const os = require('os');

// Dangerous paths that should NEVER be modified
const PROTECTED_PATHS = [
    'C:\\Windows',
    'C:\\Program Files',
    'C:\\Program Files (x86)',
    'C:\\ProgramData',
    process.env.SYSTEMROOT,
    path.join(os.homedir(), 'AppData'),
    '/System',
    '/usr',
    '/bin',
    '/etc'
].filter(Boolean);

// Dangerous commands that should be blocked
const BLOCKED_COMMANDS = [
    'format',
    'del /s',
    'rm -rf /',
    'rmdir /s',
    ':(){:|:&};:',  // Fork bomb
    'dd if=',
    'mkfs',
    '> /dev/sda',
    'chmod -R 777 /',
    'shutdown',
    'reboot',
    'reg delete',
    'bcdedit',
    'diskpart'
];

// Dangerous file extensions
const DANGEROUS_EXTENSIONS = [
    '.exe', '.bat', '.cmd', '.ps1', '.vbs', 
    '.scr', '.msi', '.dll', '.sys'
];

// Actions that require confirmation
const CONFIRM_ACTIONS = [
    'delete_file',
    'delete_folder', 
    'run_command',
    'move_file',
    'rename_file',
    // OS task actions that are destructive or irreversible
    'shutdown_computer',
    'restart_computer',
    'hibernate',
    'reboot',
    'uninstall_app',
    'kill_process',
    'end_process',
    'clear_all_cache',
    'clear_windows_update_cache',
    'write_registry',
    'delete_env_variable',
    'clear_event_log',
    'disable_scheduled_task',
    'format_drive'
];

class SafetyGuard {
    constructor() {
        this.actionLog = [];
        this.confirmationRequired = true;
        this.safeMode = true; // Enable by default
    }

    /**
     * Check if a file path is in a protected location
     */
    isProtectedPath(filePath) {
        if (!filePath) return false;
        
        const normalizedPath = path.normalize(filePath).toLowerCase();
        
        for (const protected_path of PROTECTED_PATHS) {
            if (normalizedPath.startsWith(protected_path.toLowerCase())) {
                return true;
            }
        }
        
        return false;
    }

    /**
     * Check if command contains dangerous patterns
     */
    isDangerousCommand(command) {
        if (!command) return false;
        
        const lowerCommand = command.toLowerCase();
        
        for (const blocked of BLOCKED_COMMANDS) {
            if (lowerCommand.includes(blocked.toLowerCase())) {
                return true;
            }
        }
        
        return false;
    }

    /**
     * Check if file has dangerous extension
     */
    isDangerousFile(filePath) {
        if (!filePath) return false;
        
        const ext = path.extname(filePath).toLowerCase();
        return DANGEROUS_EXTENSIONS.includes(ext);
    }

    /**
     * Check if action requires user confirmation
     */
    requiresConfirmation(actionName) {
        return this.confirmationRequired && CONFIRM_ACTIONS.includes(actionName);
    }

    /**
     * Validate an action before execution
     * Returns: { safe: boolean, reason: string, requiresConfirmation: boolean }
     */
    validateAction(actionName, params = {}) {
        const result = {
            safe: true,
            reason: '',
            requiresConfirmation: false,
            warnings: []
        };

        // Check for protected paths in file operations
        const pathParams = ['file_path', 'folder_path', 'source', 'destination', 'path', 'old_path'];
        for (const param of pathParams) {
            if (params[param] && this.isProtectedPath(params[param])) {
                result.safe = false;
                result.reason = `⛔ BLOCKED: Cannot modify protected system path: ${params[param]}`;
                return result;
            }
        }

        // Check for dangerous commands
        if (actionName === 'run_command' && params.command) {
            if (this.isDangerousCommand(params.command)) {
                result.safe = false;
                result.reason = `⛔ BLOCKED: Dangerous command detected: ${params.command}`;
                return result;
            }
            result.warnings.push('⚠️ Running system commands can be risky');
        }

        // Check for dangerous file operations
        if (actionName === 'create_file' && params.file_path) {
            if (this.isDangerousFile(params.file_path)) {
                result.safe = false;
                result.reason = `⛔ BLOCKED: Cannot create executable files: ${params.file_path}`;
                return result;
            }
        }

        // Check for delete operations
        if (actionName === 'delete_file' || actionName === 'delete_folder') {
            result.requiresConfirmation = true;
            result.warnings.push(`⚠️ This will permanently delete: ${params.file_path || params.folder_path}`);
        }

        // Warn about power/system operations
        const powerActions = ['shutdown_computer', 'restart_computer', 'hibernate', 'reboot'];
        if (powerActions.includes(actionName)) {
            result.warnings.push(`⚠️ Power action: computer will ${actionName.replace('_computer', '')}`);
        }

        // Warn about process termination
        if (actionName === 'kill_process' || actionName === 'end_process') {
            result.warnings.push(`⚠️ This will forcibly terminate process: ${params.name_or_pid || params.name || params.pid}`);
        }

        // Warn about app uninstall
        if (actionName === 'uninstall_app') {
            result.warnings.push(`⚠️ This will uninstall: ${params.app_name || params.name}`);
        }

        // Check for bulk operations (paths with wildcards)
        const targetPath = params.file_path || params.folder_path || params.path || '';
        if (targetPath.includes('*') || targetPath.includes('?')) {
            result.safe = false;
            result.reason = '⛔ BLOCKED: Wildcard operations not allowed for safety';
            return result;
        }

        // Log the action
        this.logAction(actionName, params, result);

        // Check if confirmation needed
        if (this.requiresConfirmation(actionName)) {
            result.requiresConfirmation = true;
        }

        return result;
    }

    /**
     * Log action for audit trail
     */
    logAction(actionName, params, validationResult) {
        this.actionLog.push({
            timestamp: new Date().toISOString(),
            action: actionName,
            params: params,
            result: validationResult,
        });

        // Keep only last 100 actions
        if (this.actionLog.length > 100) {
            this.actionLog = this.actionLog.slice(-100);
        }
    }

    /**
     * Get action history
     */
    getActionLog() {
        return this.actionLog;
    }

    /**
     * Generate safety report
     */
    getSafetyReport() {
        const total = this.actionLog.length;
        const blocked = this.actionLog.filter(a => !a.result.safe).length;
        const confirmed = this.actionLog.filter(a => a.result.requiresConfirmation).length;

        return {
            totalActions: total,
            blockedActions: blocked,
            confirmedActions: confirmed,
            safeMode: this.safeMode,
            lastActions: this.actionLog.slice(-10)
        };
    }

    /**
     * Toggle safe mode
     */
    setSafeMode(enabled) {
        this.safeMode = enabled;
        this.confirmationRequired = enabled;
    }
}

module.exports = new SafetyGuard();

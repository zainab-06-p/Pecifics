const { spawn } = require('child_process');

/**
 * Publisher COM Automation Module with Persistent PowerShell Session
 * Keeps Publisher application alive across multiple operations
 */

class PublisherCOM {
    constructor() {
        this.psProcess = null;
        this.isReady = false;
    }

    /**
     * Initialize persistent PowerShell session with Publisher COM object
     */
    async initializeSession() {
        if (this.psProcess) {
            return; // Already initialized
        }

        return new Promise((resolve, reject) => {
            this.psProcess = spawn('powershell.exe', ['-NoProfile', '-NoExit', '-Command', '-'], {
                stdio: ['pipe', 'pipe', 'pipe']
            });

            let outputBuffer = '';
            const readyMarker = '>>PUBLISHER_READY<<';

            this.psProcess.stdout.on('data', (data) => {
                outputBuffer += data.toString();
                if (outputBuffer.includes(readyMarker)) {
                    this.isReady = true;
                    resolve();
                }
            });

            this.psProcess.stderr.on('data', (data) => {
                console.error('Publisher PS Error:', data.toString());
            });

            // Initialize Publisher COM object
            this.psProcess.stdin.write(`
                $global:PublisherApp = New-Object -ComObject Publisher.Application
                $global:PublisherApp.Visible = $true
                Write-Output '${readyMarker}'
            \n`);
        });
    }

    /**
     * Execute command in persistent PowerShell session
     */
    async executeInSession(command, waitForOutput = true) {
        if (!this.psProcess) {
            await this.initializeSession();
        }

        return new Promise((resolve, reject) => {
            const outputMarker = `>>OUTPUT_${Date.now()}<<`;
            let output = '';
            let isCollecting = false;

            const dataHandler = (data) => {
                const text = data.toString();
                if (text.includes(outputMarker)) {
                    isCollecting = false;
                    this.psProcess.stdout.removeListener('data', dataHandler);
                    resolve(output.trim());
                } else if (isCollecting) {
                    output += text;
                } else if (text.includes('>>START<<')) {
                    isCollecting = true;
                }
            };

            if (waitForOutput) {
                this.psProcess.stdout.on('data', dataHandler);
            }

            this.psProcess.stdin.write(`
                try {
                    Write-Output '>>START<<'
                    ${command}
                    Write-Output '${outputMarker}'
                } catch {
                    Write-Output "ERROR: $($_.Exception.Message)"
                    Write-Output '${outputMarker}'
                }
            \n`);

            if (!waitForOutput) {
                setTimeout(() => resolve('OK'), 500);
            }
        });
    }

    /**
     * Close the persistent session
     */
    closeSession() {
        if (this.psProcess) {
            this.psProcess.stdin.write('$global:PublisherApp.Quit()\nexit\n');
            this.psProcess = null;
            this.isReady = false;
        }
    }

    /**
     * Create a new Publisher publication
     */
    async createPublication(templateType = 'blank') {
        const command = `
            $global:PublisherDoc = $global:PublisherApp.NewDocument()
            Write-Output "Publication created successfully"
        `;
        
        try {
            await this.executeInSession(command);
            return { success: true, message: 'Publisher publication created' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Add text to publication
     */
    async addTextBox(text, left = 100, top = 100, width = 300, height = 100) {
        const command = `
            $page = $global:PublisherDoc.Pages.Item(1)
            $textBox = $page.Shapes.AddTextbox(1, ${left}, ${top}, ${width}, ${height})
            $textBox.TextFrame.TextRange.Text = "${text}"
            Write-Output "Text box added"
        `;
        
        try {
            await this.executeInSession(command);
            return { success: true, message: 'Text box added' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Add new page to publication
     */
    async addPage() {
        const command = `
            $global:PublisherDoc.Pages.Add()
            Write-Output "Page added"
        `;
        
        try {
            await this.executeInSession(command);
            return { success: true, message: 'Page added to publication' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Save publication
     */
    async savePublication(filename = null) {
        const command = filename 
            ? `
                $fullPath = [System.IO.Path]::GetFullPath('${filename.replace(/\\/g, '\\\\')}')
                $global:PublisherDoc.SaveAs($fullPath)
                Write-Output "Publication saved to: $fullPath"
            `
            : `
                $global:PublisherDoc.Save()
                Write-Output "Publication saved"
            `;
        
        try {
            await this.executeInSession(command);
            return { success: true, message: 'Publication saved' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Check if Publisher is running
     */
    async isPublisherActive() {
        return this.isReady && this.psProcess !== null;
    }
}

module.exports = new PublisherCOM();

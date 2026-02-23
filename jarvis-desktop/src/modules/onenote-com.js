const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

/**
 * OneNote COM Automation Module
 * Uses PowerShell to control Microsoft OneNote via COM
 */

class OneNoteCOM {
    /**
     * Execute PowerShell command for OneNote automation
     */
    async executePowerShell(script) {
        try {
            const { stdout, stderr } = await execPromise(
                `powershell -NoProfile -Command "${script.replace(/"/g, '\\"')}"`,
                { encoding: 'utf8', maxBuffer: 1024 * 1024 }
            );
            
            if (stderr && !stderr.includes('Warning')) {
                throw new Error(stderr);
            }
            
            return stdout.trim();
        } catch (error) {
            throw new Error(`OneNote COM Error: ${error.message}`);
        }
    }

    /**
     * Open OneNote
     */
    async openOneNote() {
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            Start-Process "onenote:"
            Start-Sleep -Seconds 2
            Write-Output "OneNote opened"
        `;
        
        try {
            await this.executePowerShell(script);
            return { success: true, message: 'OneNote opened' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Create a new page with title
     */
    async createPage(title, sectionName = null) {
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            
            # Get current section
            [ref]$currentSectionId = ""
            $onenote.GetCurrentPage([ref]$currentSectionId)
            
            # Create new page
            [ref]$newPageId = ""
            $onenote.CreateNewPage($currentSectionId.Value, [ref]$newPageId)
            
            # Get page content
            [ref]$pageXml = ""
            $onenote.GetPageContent($newPageId.Value, [ref]$pageXml)
            
            # Update title
            $xml = [xml]$pageXml.Value
            $xml.Page.Title.OE.T = "${title}"
            
            # Update page
            $onenote.UpdatePageContent($xml.OuterXml)
            
            Write-Output "Page created with title: ${title}"
        `;
        
        try {
            await this.executePowerShell(script);
            return { success: true, message: `Page '${title}' created` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Add content to current page
     */
    async addContent(content) {
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            
            # Get current page
            [ref]$pageId = ""
            $onenote.GetCurrentPage([ref]$pageId)
            
            # Get page content
            [ref]$pageXml = ""
            $onenote.GetPageContent($pageId.Value, [ref]$pageXml)
            
            # Add content (simplified - adds text outline)
            $xml = [xml]$pageXml.Value
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("one", $xml.DocumentElement.NamespaceURI)
            
            # Create new outline
            $outline = $xml.CreateElement("one", "Outline", $xml.DocumentElement.NamespaceURI)
            $oeChildren = $xml.CreateElement("one", "OEChildren", $xml.DocumentElement.NamespaceURI)
            $oe = $xml.CreateElement("one", "OE", $xml.DocumentElement.NamespaceURI)
            $t = $xml.CreateElement("one", "T", $xml.DocumentElement.NamespaceURI)
            $t.InnerText = "${content}"
            
            $oe.AppendChild($t)
            $oeChildren.AppendChild($oe)
            $outline.AppendChild($oeChildren)
            $xml.DocumentElement.AppendChild($outline)
            
            # Update page
            $onenote.UpdatePageContent($xml.OuterXml)
            
            Write-Output "Content added"
        `;
        
        try {
            await this.executePowerShell(script);
            return { success: true, message: 'Content added to OneNote' };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Check if OneNote is running
     */
    async isOneNoteActive() {
        const script = `
            try {
                $onenote = New-Object -ComObject OneNote.Application
                Write-Output "true"
            } catch {
                Write-Output "false"
            }
        `;
        
        try {
            const result = await this.executePowerShell(script);
            return result === 'true';
        } catch {
            return false;
        }
    }
}

module.exports = new OneNoteCOM();

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
        // Use GetHierarchy to find the current notebook's first section,
        // then create a page in it.
        const escapedTitle = (title || 'New Page').replace(/"/g, '`"').replace(/'/g, "''");
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            
            # Get hierarchy XML to find first section
            [ref]$hierarchyXml = ""
            $onenote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$hierarchyXml)
            
            $xml = [xml]$hierarchyXml.Value
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("one", $xml.DocumentElement.NamespaceURI)
            
            # Find first section
            $section = $xml.SelectSingleNode("//one:Section", $ns)
            if (-not $section) {
                Write-Output "ERROR:No section found in OneNote"
                return
            }
            $sectionId = $section.GetAttribute("ID")
            
            # Create new page
            [ref]$newPageId = ""
            $onenote.CreateNewPage($sectionId, [ref]$newPageId)
            
            # Get the new page's XML
            [ref]$pageXml = ""
            $onenote.GetPageContent($newPageId.Value, [ref]$pageXml)
            
            # Set title
            $pageDoc = [xml]$pageXml.Value
            $nsPage = New-Object System.Xml.XmlNamespaceManager($pageDoc.NameTable)
            $nsPage.AddNamespace("one", $pageDoc.DocumentElement.NamespaceURI)
            $titleNode = $pageDoc.SelectSingleNode("//one:Title/one:OE/one:T", $nsPage)
            if ($titleNode) {
                $titleNode.InnerText = '${escapedTitle}'
            }
            $onenote.UpdatePageContent($pageDoc.OuterXml)
            
            # Navigate to the new page
            $onenote.NavigateTo($newPageId.Value)
            Write-Output "Page created: ${escapedTitle}"
        `;
        
        try {
            const result = await this.executePowerShell(script);
            if (result.startsWith('ERROR:')) {
                return { success: false, error: result.substring(6) };
            }
            return { success: true, message: `Page '${title}' created in OneNote` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    /**
     * Add content to current page
     */
    async addContent(content) {
        const escapedContent = (content || '').replace(/"/g, '`"').replace(/'/g, "''").replace(/\n/g, '<br/>');
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            
            # Get the current page ID
            [ref]$currentPageId = ""
            [ref]$pageXml = ""
            $onenote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages, [ref]$pageXml)
            
            # Get the currently active page
            $xml = [xml]$pageXml.Value
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("one", $xml.DocumentElement.NamespaceURI)
            $activePage = $xml.SelectSingleNode("//one:Page[@isCurrentlyViewed='true']", $ns)
            
            if (-not $activePage) {
                # Fall back to first page
                $activePage = $xml.SelectSingleNode("//one:Page", $ns)
            }
            
            if (-not $activePage) {
                Write-Output "ERROR:No page found"
                return
            }
            
            $pageId = $activePage.GetAttribute("ID")
            
            # Get full page content
            [ref]$fullPageXml = ""
            $onenote.GetPageContent($pageId, [ref]$fullPageXml)
            $pageDoc = [xml]$fullPageXml.Value
            $nsP = New-Object System.Xml.XmlNamespaceManager($pageDoc.NameTable)
            $nsP.AddNamespace("one", $pageDoc.DocumentElement.NamespaceURI)
            
            # Create new outline with the content
            $oneNs = $pageDoc.DocumentElement.NamespaceURI
            $outline = $pageDoc.CreateElement("one", "Outline", $oneNs)
            $oeChildren = $pageDoc.CreateElement("one", "OEChildren", $oneNs)
            $oe = $pageDoc.CreateElement("one", "OE", $oneNs)
            $t = $pageDoc.CreateElement("one", "T", $oneNs)
            $cdata = $pageDoc.CreateCDataSection('${escapedContent}')
            $t.AppendChild($cdata)
            $oe.AppendChild($t)
            $oeChildren.AppendChild($oe)
            $outline.AppendChild($oeChildren)
            $pageDoc.DocumentElement.AppendChild($outline)
            
            # Update page
            $onenote.UpdatePageContent($pageDoc.OuterXml)
            Write-Output "Content added"
        `;
        
        try {
            const result = await this.executePowerShell(script);
            if (result.startsWith('ERROR:')) {
                return { success: false, error: result.substring(6) };
            }
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

    /**
     * List all notebooks and their sections
     */
    async listNotebooks() {
        const script = `
            $onenote = New-Object -ComObject OneNote.Application
            [ref]$xml = ""
            $onenote.GetHierarchy("", [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsSections, [ref]$xml)
            $doc = [xml]$xml.Value
            $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
            $ns.AddNamespace("one", $doc.DocumentElement.NamespaceURI)
            $results = @()
            foreach ($nb in $doc.SelectNodes("//one:Notebook", $ns)) {
                $sections = @()
                foreach ($sec in $nb.SelectNodes("one:Section", $ns)) {
                    $sections += $sec.GetAttribute("name")
                }
                $results += [PSCustomObject]@{ name = $nb.GetAttribute("name"); sections = ($sections -join ", ") }
            }
            $results | ConvertTo-Json -Compress
        `;
        try {
            const result = await this.executePowerShell(script);
            let notebooks = JSON.parse(result);
            if (!Array.isArray(notebooks)) notebooks = [notebooks];
            return { success: true, notebooks, message: `Found ${notebooks.length} notebook(s)` };
        } catch (error) {
            return { success: false, error: error.message };
        }
    }
}

module.exports = new OneNoteCOM();

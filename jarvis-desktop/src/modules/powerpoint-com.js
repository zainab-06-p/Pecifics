const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

/**
 * PowerPoint COM Automation Module
 * Uses PowerShell to control PowerPoint via COM
 */

class PowerPointCOM {
    /**
     * Execute PowerShell command for PowerPoint automation
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
            throw new Error(`PowerPoint COM Error: ${error.message}`);
        }
    }

    /**
     * Apply design theme to active PowerPoint presentation
     */
    async applyTheme(themeName) {
        const themeMap = {
            'ion': 'Ion',
            'ion_boardroom': 'Ion Boardroom',
            'facet': 'Facet',
            'integral': 'Integral',
            'office_theme': 'Office Theme',
            'slice': 'Slice',
            'wisp': 'Wisp',
            'organic': 'Organic',
            'retrospect': 'Retrospect',
            'dividend': 'Dividend',
            'basis': 'Basis',
            'berlin': 'Berlin',
            'circuit': 'Circuit'
        };

        const displayName = themeMap[themeName.toLowerCase()] || themeName;

        const script = `
            try {
                $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');
                $pres = $ppt.ActivePresentation;
            
            # Try to apply theme by name
            try {
                $themeFound = $false;
                foreach ($design in $pres.Designs) {
                    if ($design.Name -like '*${displayName}*') {
                        $pres.ApplyTheme($design.ThemePath);
                        $themeFound = $true;
                        break;
                    }
                }
                if (-not $themeFound) {
                    # Try built-in themes path
                    $themePath = "$env:ProgramFiles\\Microsoft Office\\root\\Document Themes 16\\${displayName}.thmx";
                    if (Test-Path $themePath) {
                        $pres.ApplyTheme($themePath);
                    } else {
                        Write-Host "Theme '${displayName}' not found, using default";
                    }
                }
                Write-Host "Theme applied successfully";
            } catch {
                Write-Host "Could not apply theme: $_";
            }
        `;

        return await this.executePowerShell(script);
    }

    /**
     * Add animation to current slide
     */
    async addAnimation(animationType, applyToAll = false) {
        const animationMap = {
            'fade': 'ppEffectFade',
            'fly_in': 'ppEffectFly',
            'wipe': 'ppEffectWipe',
            'split': 'ppEffectSplit',
            'appear': 'ppEffectAppear',
            'zoom': 'ppEffectZoom',
            'swivel': 'ppEffectSwivel',
            'bounce': 'ppEffectBounce'
        };

        const effectName = animationMap[animationType.toLowerCase()] || 'ppEffectFade';
        const effectValue = this.getAnimationValue(effectName);

        const script = `
            try {
                $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');
                $pres = $ppt.ActivePresentation;
            
            ${applyToAll ? `
                # Apply to all slides
                foreach ($slide in $pres.Slides) {
                    if ($slide.Shapes.Count -gt 0) {
                        $shape = $slide.Shapes.Item(1);
                        $effect = $slide.TimeLine.MainSequence.AddEffect($shape, ${effectValue}, 1, 1);
                    }
                }
                Write-Host "Animation applied to all slides";
            ` : `
                # Apply to current slide only
                $slide = $ppt.ActiveWindow.View.Slide;
                if ($slide.Shapes.Count -gt 0) {
                    $shape = $slide.Shapes.Item(1);
                    $effect = $slide.TimeLine.MainSequence.AddEffect($shape, ${effectValue}, 1, 1);
                }
                Write-Host "Animation applied to current slide";
            `}
        `;

        return await this.executePowerShell(script);
    }

    /**
     * Change layout of current slide
     */
    async changeLayout(layoutName) {
        const layoutMap = {
            'title_slide': 1,
            'title_and_content': 2,
            'section_header': 3,
            'two_content': 4,
            'comparison': 5,
            'title_only': 6,
            'blank': 7,
            'content_with_caption': 8,
            'picture_with_caption': 9
        };

        const layoutIndex = layoutMap[layoutName.toLowerCase()] || 2;

        const script = `
            $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');
            if ($ppt.Presentations.Count -eq 0) { 
                throw 'No active PowerPoint presentation found';
            }
            $pres = $ppt.ActivePresentation;
            $slide = $ppt.ActiveWindow.View.Slide;
            
            try {
                $layout = $pres.SlideMaster.CustomLayouts.Item(${layoutIndex});
                $slide.CustomLayout = $layout;
                Write-Host "Layout changed successfully";
            } catch {
                Write-Host "Could not change layout: $_";
            }
        `;

        return await this.executePowerShell(script);
    }

    /**
     * Get PowerPoint animation constant value
     */
    getAnimationValue(effectName) {
        const constants = {
            'ppEffectFade': 10,
            'ppEffectFly': 2,
            'ppEffectWipe': 3,
            'ppEffectSplit': 40,
            'ppEffectAppear': 1,
            'ppEffectZoom': 16,
            'ppEffectSwivel': 41,
            'ppEffectBounce': 30
        };
        return constants[effectName] || 10;
    }

    /**
     * Find which slides contain a given text string
     * Returns array of { SlideNumber, ShapeName, Text } matches
     */
    async findSlideByText(searchText) {
        const esc = searchText.replace(/'/g, "''");
        const script = `
            try {
                $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
                $pres = $ppt.ActivePresentation
                $results = @()
                for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                    $slide = $pres.Slides.Item($i)
                    foreach ($shape in $slide.Shapes) {
                        if ($shape.HasTextFrame) {
                            $txt = $shape.TextFrame.TextRange.Text
                            if ($txt -like '*${esc}*') {
                                $results += [PSCustomObject]@{
                                    SlideNumber = $i
                                    ShapeName   = $shape.Name
                                    Text        = $txt.Substring(0, [Math]::Min(200, $txt.Length))
                                }
                            }
                        }
                    }
                }
                if ($results.Count -eq 0) { Write-Host '[]' }
                else { $results | ConvertTo-Json -Compress | Write-Host }
            } catch { Write-Host "ERROR: $_" }
        `;
        const raw = await this.executePowerShell(script);
        try {
            const jsonMatch = raw.match(/(\[.*\]|\{.*\})/s);
            const data = JSON.parse(jsonMatch ? jsonMatch[0] : '[]');
            return { success: true, slides: Array.isArray(data) ? data : [data] };
        } catch {
            return { success: false, error: raw };
        }
    }

    /**
     * Get all text content from a specific slide
     * Returns array of { Shape, Text } objects
     */
    async getSlideContent(slideNumber) {
        const script = `
            try {
                $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
                $pres = $ppt.ActivePresentation
                $slide = $pres.Slides.Item(${slideNumber})
                $shapes = @()
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame) {
                        $shapes += [PSCustomObject]@{ Shape = $shape.Name; Text = $shape.TextFrame.TextRange.Text }
                    }
                }
                if ($shapes.Count -eq 0) { Write-Host '[]' }
                else { $shapes | ConvertTo-Json -Compress | Write-Host }
            } catch { Write-Host "ERROR: $_" }
        `;
        const raw = await this.executePowerShell(script);
        try {
            const jsonMatch = raw.match(/(\[.*\]|\{.*\})/s);
            const data = JSON.parse(jsonMatch ? jsonMatch[0] : '[]');
            return { success: true, shapes: Array.isArray(data) ? data : [data] };
        } catch {
            return { success: false, error: raw };
        }
    }

    /**
     * Replace text on a specific slide (all matching shapes)
     * @param {number} slideNumber - 1-based slide index
     * @param {string} oldText - Text to find
     * @param {string} newText - Replacement text
     */
    async updateSlideText(slideNumber, oldText, newText) {
        const escOld = oldText.replace(/'/g, "''").replace(/"/g, '`"');
        const escNew = newText.replace(/'/g, "''").replace(/"/g, '`"');
        const script = `
            try {
                $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
                $pres = $ppt.ActivePresentation
                $slide = $pres.Slides.Item(${slideNumber})
                $count = 0
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame) {
                        $range = $shape.TextFrame.TextRange
                        if ($range.Text -like '*${escOld}*') {
                            $range.Text = $range.Text.Replace("${escOld}", "${escNew}")
                            $count++
                        }
                    }
                }
                $pres.Save()
                Write-Host "Updated $count shape(s) on slide ${slideNumber}"
            } catch { Write-Host "ERROR: $_" }
        `;
        const raw = await this.executePowerShell(script);
        return { success: !raw.includes('ERROR'), message: raw.trim() };
    }

    /**
     * Check if PowerPoint is running with active presentation
     */
    async isPowerPointActive() {
        try {
            const script = `
                try {
                    $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application');
                    if ($ppt.Presentations.Count -gt 0) {
                        Write-Host "true";
                    } else {
                        Write-Host "false";
                    }
                } catch {
                    Write-Host "false";
                }
            `;
            const result = await this.executePowerShell(script);
            return result.includes('true');
        } catch {
            return false;
        }
    }
}

module.exports = new PowerPointCOM();

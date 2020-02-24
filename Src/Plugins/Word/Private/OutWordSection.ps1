function OutWordSection
{
<#
    .SYNOPSIS
        Output formatted Word section (paragraph).
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Section,

        [Parameter(Mandatory)]
        [System.Xml.XmlElement] $RootElement,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $p = $RootElement.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));

        if (-not [System.String]::IsNullOrEmpty($Section.Style))
        {
            #if (-not $Section.IsExcluded) {
            ## If it's excluded we need a non-Heading style :( Could explicitly set the style on the run?
            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
            [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Section.Style)
            #}
        }

        if ($Section.Tabs -gt 0)
        {
            $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain));
            [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Section.Tabs));
        }

        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
        ## Increment heading spacing by 2pt for each section level, starting at 8pt for level 0, 10pt for level 1 etc
        $spacingPt = (($Section.Level * 2) + 8) * 20
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, $spacingPt)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, $spacingPt)
        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))

        if ($Document.Options['EnableSectionNumbering'])
        {
            [System.String] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name
        }
        else
        {
            [System.String] $sectionName = '{0}' -f $Section.Name
        }

        [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($sectionName))

        foreach ($s in $Section.Sections.GetEnumerator())
        {
            if ($s.Id.Length -gt 40)
            {
                $sectionId = '{0}[..]' -f $s.Id.Substring(0, 36)
            }
            else
            {
                $sectionId = $s.Id
            }
            $currentIndentationLevel = 1
            if ($null -ne $s.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $s.Level + 1
            }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel
            switch ($s.Type)
            {
                'PScribo.Section' {
                    $s | OutWordSection -RootElement $RootElement -XmlDocument $XmlDocument
                }
                'PScribo.Paragraph' {
                    [ref] $null = $RootElement.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $XmlDocument))
                }
                'PScribo.PageBreak' {
                    [ref] $null = $RootElement.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $XmlDocument))
                }
                'PScribo.LineBreak' {
                    [ref] $null = $RootElement.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $XmlDocument))
                }
                'PScribo.Table' {
                    OutWordTable -Table $s -XmlDocument $XmlDocument -Element $RootElement
                }
                'PScribo.BlankLine' {
                    OutWordBlankLine -BlankLine $s -XmlDocument $XmlDocument -Element $RootElement
                }
                'PScribo.Image' {
                    [ref] $null = $RootElement.AppendChild((OutWordImage -Image $s -XmlDocument $XmlDocument))
                }
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning
                }
            }
        }

        if ($Section.IsSectionBreakEnd)
        {
            $sectionPrParams = @{
                PageHeight       = if ($Section.Orientation -eq 'Portrait') { $Document.Options['PageHeight'] } else { $Document.Options['PageWidth'] }
                PageWidth        = if ($Section.Orientation -eq 'Portrait') { $Document.Options['PageWidth'] } else { $Document.Options['PageHeight'] }
                PageMarginTop    = $Document.Options['MarginTop'];
                PageMarginBottom = $Document.Options['MarginBottom'];
                PageMarginLeft   = $Document.Options['MarginLeft'];
                PageMarginRight  = $Document.Options['MarginRight'];
                Orientation      = $Section.Orientation;
            }
            [ref] $null = $pPr.AppendChild((GetWordSectionPr @sectionPrParams -XmlDocument $xmlDocument));
        }
    }
}

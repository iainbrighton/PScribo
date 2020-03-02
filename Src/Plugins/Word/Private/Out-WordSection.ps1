function Out-WordSection
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
        [System.Xml.XmlElement] $Element,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $p = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));

        if (-not [System.String]::IsNullOrEmpty($Section.Style))
        {
            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
            [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Section.Style)
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

        if ($Section.IsSectionBreakEnd)
        {
            $sectionPrParams = @{
                PageHeight       = $Document.Options['PageHeight']
                PageWidth        = $Document.Options['PageWidth']
                PageMarginTop    = $Document.Options['MarginTop']
                PageMarginBottom = $Document.Options['MarginBottom']
                PageMarginLeft   = $Document.Options['MarginLeft']
                PageMarginRight  = $Document.Options['MarginRight']
                Orientation      = $Section.Orientation;
            }
            [ref] $null = $pPr.AppendChild((Get-WordSectionPr @sectionPrParams -XmlDocument $xmlDocument))
        }


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

        foreach ($subSection in $Section.Sections.GetEnumerator())
        {
            if ($subSection.Id.Length -gt 40)
            {
                $sectionId = '{0}[..]' -f $subSection.Id.Substring(0, 36)
            }
            else
            {
                $sectionId = $subSection.Id
            }

            $currentIndentationLevel = 1
            if ($null -ne $subSection.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $subSection.Level + 1
            }
            WriteLog -Message ($localized.PluginProcessingSection -f $subSection.Type, $sectionId) -Indent $currentIndentationLevel

            switch ($subSection.Type)
            {
                'PScribo.Section' {
                    Out-WordSection -Section $subSection -Element $Element -XmlDocument $XmlDocument
                }
                'PScribo.Paragraph' {
                    [ref] $null = $Element.AppendChild((Out-WordParagraph -Paragraph $subSection -XmlDocument $XmlDocument))
                }
                'PScribo.PageBreak' {
                    [ref] $null = $Element.AppendChild((Out-WordPageBreak -PageBreak $subSection -XmlDocument $XmlDocument))
                }
                'PScribo.LineBreak' {
                    [ref] $null = $Element.AppendChild((Out-WordLineBreak -LineBreak $subSection -XmlDocument $XmlDocument))
                }
                'PScribo.Table' {
                    Out-WordTable -Table $subSection -XmlDocument $XmlDocument -Element $Element
                }
                'PScribo.BlankLine' {
                    Out-WordBlankLine -BlankLine $subSection -XmlDocument $XmlDocument -Element $Element
                }
                'PScribo.Image' {
                    [ref] $null = $Element.AppendChild((Out-WordImage -Image $subSection -XmlDocument $XmlDocument))
                }
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }
    }
}

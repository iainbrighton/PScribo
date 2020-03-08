function Out-WordParagraph
{
<#
    .SYNOPSIS
        Output formatted Word paragraph.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Paragraph,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    begin
    {
        if ([System.String]::IsNullOrEmpty($Paragraph.Text))
        {
            $lines = $Paragraph.Id.Split([System.Environment]::NewLine)
        }
        else
        {
            $lines = $Paragraph.Text.Split([System.Environment]::NewLine)
        }
    }
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $p = $XmlDocument.CreateElement('w', 'p', $xmlns);
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns));

        if ($Paragraph.Tabs -gt 0)
        {
            $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlns))
            [ref] $null = $ind.SetAttribute('left', $xmlns, (720 * $Paragraph.Tabs))
        }
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlns))
            [ref] $null = $pStyle.SetAttribute('val', $xmlns, $Paragraph.Style)
        }

        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('before', $xmlns, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlns, 0)

        if ($Paragraph.IsSectionBreakEnd)
        {
            $paragraphPrParams = @{
                PageHeight       = $Document.Options['PageHeight']
                PageWidth        = $Document.Options['PageWidth']
                PageMarginTop    = $Document.Options['MarginTop'];
                PageMarginBottom = $Document.Options['MarginBottom'];
                PageMarginLeft   = $Document.Options['MarginLeft'];
                PageMarginRight  = $Document.Options['MarginRight'];
                Orientation      = $Paragraph.Orientation;
            }
            [ref] $null = $pPr.AppendChild((Get-WordSectionPr @paragraphPrParams -XmlDocument $xmlDocument))
        }

        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlns))

        if ($Paragraph.Font)
        {
            $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlns))
            [ref] $null = $rFonts.SetAttribute('ascii', $xmlns, $Paragraph.Font[0])
            [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlns, $Paragraph.Font[0])
        }

        if ($Paragraph.Size -gt 0)
        {
            $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlns))
            [ref] $null = $sz.SetAttribute('val', $xmlns, $Paragraph.Size * 2)
        }

        if ($Paragraph.Bold -eq $true)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlns))
        }

        if ($Paragraph.Italic -eq $true)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlns))
        }

        if ($Paragraph.Underline -eq $true)
        {
            $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlns))
            [ref] $null = $u.SetAttribute('val', $xmlns, 'single')
        }

        if (-not [System.String]::IsNullOrEmpty($Paragraph.Color))
        {
            $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlns))
            [ref] $null = $Color.SetAttribute('val', $xmlns, (ConvertTo-WordColor -Color $Paragraph.Color))
        }

        ## Create a separate run for each line/break
        for ($l = 0; $l -lt $lines.Count; $l++)
        {
            $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
            ## needs to be xml:space="preserve" NOT w:space...
            [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
            [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]))

            if ($l -lt ($lines.Count - 1))
            {
                ## Don't add a line break to the last line/break
                [ref] $null = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlns))
            }
        }

        return $p;
    }
}

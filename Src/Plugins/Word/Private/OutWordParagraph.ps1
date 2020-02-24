function OutWordParagraph
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
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));

        if ($Paragraph.Tabs -gt 0)
        {
            $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain))
            [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Paragraph.Tabs))
        }
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
            [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Paragraph.Style)
        }

        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)

        if ([System.String]::IsNullOrEmpty($Paragraph.Text))
        {
            $lines = $Paragraph.Id -Split [System.Environment]::NewLine
        }
        else
        {
            $lines = $Paragraph.TexT -Split [System.Environment]::NewLine
        }

        ## Create a separate run for each line/break
        for ($l = 0; $l -lt $lines.Count; $l++)
        {
            $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
            $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain))
            ## Apply custom paragraph styles to the run..
            if ($Paragraph.Font)
            {
                $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain))
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Paragraph.Font[0])
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Paragraph.Font[0])
            }
            if ($Paragraph.Size -gt 0)
            {
                $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain))
                [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Paragraph.Size * 2)
            }
            if ($Paragraph.Bold -eq $true)
            {
                [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain))
            }
            if ($Paragraph.Italic -eq $true)
            {
                [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain))
            }
            if ($Paragraph.Underline -eq $true)
            {
                $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain))
                [ref] $null = $u.SetAttribute('val', $xmlnsMain, 'single')
            }
            if (-not [System.String]::IsNullOrEmpty($Paragraph.Color))
            {
                $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
                [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Paragraph.Color))
            }

            $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
            [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
            ## needs to be xml:space="preserve" NOT w:space...
            [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]))

            if ($l -lt ($lines.Count - 1))
            {
                ## Don't add a line break to the last line/break
                $brr = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $brt = $brr.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                [ref] $null = $brt.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain))
            }
        }

        if ($Paragraph.IsSectionBreakEnd)
        {
            $paragraphPrParams = @{
                PageHeight       = if ($Paragraph.Orientation -eq 'Portrait') { $Document.Options['PageHeight'] } else { $Document.Options['PageWidth'] }
                PageWidth        = if ($Paragraph.Orientation -eq 'Portrait') { $Document.Options['PageWidth'] } else { $Document.Options['PageHeight'] }
                PageMarginTop    = $Document.Options['MarginTop'];
                PageMarginBottom = $Document.Options['MarginBottom'];
                PageMarginLeft   = $Document.Options['MarginLeft'];
                PageMarginRight  = $Document.Options['MarginRight'];
                Orientation      = $Paragraph.Orientation;
            }
            [ref] $null = $pPr.AppendChild((GetWordSectionPr @paragraphPrParams -XmlDocument $xmlDocument));
        }
        return $p;
    }
}

function Get-WordParagraphRunPr
{
<#
    .SYNOPSIS
        Outputs paragraph text/run (rPr) properties.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.PSObject] $ParagraphRun,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlns)

        if ($ParagraphRun.HasStyle -eq $true)
        {
            $rStyle = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rStyle', $xmlns))
            $characterStyle = '{0}Char' -f $ParagraphRun.Style
            [ref] $null = $rStyle.SetAttribute('val', $xmlns, $characterStyle)
        }
        if ($ParagraphRun.HasInlineStyle -eq $true)
        {
            if ($null -ne $ParagraphRun.Font)
            {
                $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlns))
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlns, $ParagraphRun.Font[0])
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlns, $ParagraphRun.Font[0])
            }

            if ($ParagraphRun.Bold -eq $true)
            {
                [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlns))
            }

            if ($ParagraphRun.Underline -eq $true)
            {
                $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlns))
                [ref] $null = $u.SetAttribute('val', $xmlns, 'single')
            }

            if ($ParagraphRun.Italic -eq $true)
            {
                [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlns))
            }

            if (-not [System.String]::IsNullOrEmpty($ParagraphRun.Color))
            {
                $wordColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $ParagraphRun.Color)
                $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlns))
                [ref] $null = $color.SetAttribute('val', $xmlns, $wordColor)
            }

            if (($null -ne $ParagraphRun.Size) -and ($ParagraphRun.Size -gt 0))
            {
                $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlns))
                [ref] $null = $sz.SetAttribute('val', $xmlns, $ParagraphRun.Size * 2)
            }
        }

        return $rPr
    }
}

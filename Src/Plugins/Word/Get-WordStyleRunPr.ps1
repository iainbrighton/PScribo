function Get-WordStyleRunPr
{
<#
    .SYNOPSIS
        Generates Word run (rPr) formatting properties
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSObject] $Style,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlns)

        $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlns))
        [ref] $null = $rFonts.SetAttribute('ascii', $xmlns, $Style.Font[0])
        [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlns, $Style.Font[0])

        if ($Style.Bold)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlns))
        }

        if ($Style.Underline)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlns))
        }

        if ($Style.Italic)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlns))
        }

        $wordColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $Style.Color)
        $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlns))
        [ref] $null = $color.SetAttribute('val', $xmlns, $wordColor)

        $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlns))
        [ref] $null = $sz.SetAttribute('val', $xmlns, $Style.Size * 2)

        return $rPr
    }
}

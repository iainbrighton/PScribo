function GetWordStyleRunPr
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
        $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain)
        $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain))
        [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Style.Font[0])
        [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Style.Font[0])

        if ($Style.Bold)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain))
        }
        if ($Style.Underline)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain))
        }
        if ($Style.Italic)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain))
        }

        $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
        [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color))
        $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain))
        [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Style.Size * 2)
        return $rPr
    }
}

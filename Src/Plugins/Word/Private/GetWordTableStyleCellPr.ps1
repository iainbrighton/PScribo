function GetWordTableStyleCellPr
{
<#
    .SYNOPSIS
        Generates Word table cell (tcPr) formatting properties
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
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tcPr = $XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain)

        if ($Style.BackgroundColor)
        {
            $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain))
            [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear')
            [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto')
            [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor))
        }
        return $tcPr
    }
}

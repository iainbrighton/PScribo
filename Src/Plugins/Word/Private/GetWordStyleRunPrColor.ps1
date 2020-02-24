function GetWordStyleRunPrColor
{
<#
    .SYNOPSIS
        Generates Word run (rPr) text colour formatting property only.

    .NOTES
        This is only required to override the text colour in table rows/headers
        as I can't get this (yet) applied via the table style?
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
        $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
        [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color))
        return $rPr
    }
}

function OutWordPageBreak
{
<#
    .SYNOPSIS
    Output formatted Word page break.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $PageBreak,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        $br = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain))
        [ref] $null = $br.SetAttribute('type', $xmlnsMain, 'page')
        return $p
    }
}

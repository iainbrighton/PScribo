function Out-WordLineBreak
{
<#
    .SYNOPSIS
    Output formatted Word line break.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $LineBreak,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
        $pBdr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pBdr', $xmlnsMain))
        $bottom = $pBdr.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain))
        [ref] $null = $bottom.SetAttribute('val', $xmlnsMain, 'single')
        [ref] $null = $bottom.SetAttribute('sz', $xmlnsMain, 6)
        [ref] $null = $bottom.SetAttribute('space', $xmlnsMain, 1)
        [ref] $null = $bottom.SetAttribute('color', $xmlnsMain, 'auto')
        return $p
    }
}

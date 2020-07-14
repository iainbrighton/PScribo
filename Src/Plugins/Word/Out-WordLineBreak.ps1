function Out-WordLineBreak
{
<#
    .SYNOPSIS
        Output formatted Word line break.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','LineBreak')]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $LineBreak,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlns)
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        $pBdr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pBdr', $xmlns))
        $bottom = $pBdr.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlns))
        [ref] $null = $bottom.SetAttribute('val', $xmlns, 'single')
        [ref] $null = $bottom.SetAttribute('sz', $xmlns, 6)
        [ref] $null = $bottom.SetAttribute('space', $xmlns, 1)
        [ref] $null = $bottom.SetAttribute('color', $xmlns, 'auto')

        return $p
    }
}

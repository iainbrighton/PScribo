function Out-WordPageBreak
{
<#
    .SYNOPSIS
        Output formatted Word page break.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','PageBreak')]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $PageBreak,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlns)
        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $br = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlns))
        [ref] $null = $br.SetAttribute('type', $xmlns, 'page')

        return $p
    }
}

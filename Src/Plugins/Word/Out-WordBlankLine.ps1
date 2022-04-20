function Out-WordBlankLine
{
<#
    .SYNOPSIS
        Output formatted Word xml blank line (empty paragraph).
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $BlankLine,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument,

        [Parameter(Mandatory)]
        [System.Xml.XmlElement] $Element
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
            [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        }
    }
}

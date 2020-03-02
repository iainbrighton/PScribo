function Get-WordStyleParagraphPr2222
{
<#
    .SYNOPSIS
        Generates Word paragraph (pPr) formatting properties
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

        $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlns)
        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('before', $xmlns, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlns, 0)

        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlns))
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlns))
        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlns))

        if ($Style.Align.ToLower() -eq 'justify')
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, 'distribute')
        }
        else
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, $Style.Align.ToLower())
        }

        return $pPr
    }
}

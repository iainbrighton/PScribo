function GetWordStyleParagraphPr
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
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlnsMain)
        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain))
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain))
        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))

        if ($Style.Align.ToLower() -eq 'justify')
        {
            [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
        }
        else
        {
            [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower())
        }
        return $pPr
    }
}

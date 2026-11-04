function Get-WordNumberingDocument
{
<#
    .SYNOPSIS
        Outputs Office Open XML numbering document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param
    (
        ## PScribo document styles
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList] $Lists,

        ## Adds the multilevel list that numbers the section headings
        [Parameter()]
        [System.Management.Automation.SwitchParameter] $EnableWordSectionNumbering
    )
    process
    {
        ## Create the numbering.xml document
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $xmlDocument.AppendChild($XmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $numbering = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'numbering', $xmlns))
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:w14', 'http://schemas.microsoft.com/office/word/2010/wordml')

        foreach ($list in $Lists.GetEnumerator())
        {
            $abstractNum = Get-WordNumberStyle -List $list -XmlDocument $xmlDocument
            [ref] $null = $numbering.AppendChild($abstractNum)
        }

        $sectionNumberId = Get-WordSectionNumberId -Lists $Lists
        ## All w:abstractNum elements have to precede the w:num elements
        if ($EnableWordSectionNumbering)
        {
            $abstractNum = Get-WordSectionNumberStyle -AbstractNumberId ($sectionNumberId -1) -XmlDocument $xmlDocument
            [ref] $null = $numbering.AppendChild($abstractNum)
        }

        foreach ($list in $Lists.GetEnumerator())
        {
            $num = $numbering.AppendChild($xmlDocument.CreateElement('w', 'num', $xmlns))
            [ref] $null = $num.SetAttribute('numId', $xmlns, $list.Number)
            $abstractNumId = $num.AppendChild($xmlDocument.CreateElement('w', 'abstractNumId', $xmlns))
            [ref] $null = $abstractNumId.SetAttribute('val', $xmlns, $list.Number -1)
        }

        if ($EnableWordSectionNumbering)
        {
            $num = $numbering.AppendChild($xmlDocument.CreateElement('w', 'num', $xmlns))
            [ref] $null = $num.SetAttribute('numId', $xmlns, $sectionNumberId)
            $abstractNumId = $num.AppendChild($xmlDocument.CreateElement('w', 'abstractNumId', $xmlns))
            [ref] $null = $abstractNumId.SetAttribute('val', $xmlns, $sectionNumberId -1)
        }

        return $xmlDocument
    }
}

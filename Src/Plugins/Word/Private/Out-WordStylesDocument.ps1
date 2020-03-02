function Out-WordStylesDocument
{
<#
    .SYNOPSIS
        Outputs Office Open XML style document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param
    (
        ## PScribo document styles
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Collections.Hashtable] $Styles,

        ## PScribo document tables styles
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Collections.Hashtable] $TableStyles
    )
    process
    {
        ## Create the Style.xml document
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $xmlDocument.AppendChild($XmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $documentStyles = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'styles', $xmlnsMain))

        ## Create default style
        $defaultStyle = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'style', $xmlnsMain))
        [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'paragraph')
        [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1')
        [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'Normal')
        $defaultStyleName = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlnsMain))
        [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'Normal')
        [ref] $null = $defaultStyle.AppendChild($xmlDocument.CreateElement('w', 'qFormat', $xmlnsMain))

        foreach ($Style in $Styles.Values)
        {
            $documentParagraphStyle = Get-WordStyle -Style $Style -XmlDocument $xmlDocument -Type Paragraph
            [ref] $null = $documentStyles.AppendChild($documentParagraphStyle)
            $documentCharacterStyle = Get-WordStyle -Style $Style -XmlDocument $xmlDocument -Type Character
            [ref] $null = $documentStyles.AppendChild($documentCharacterStyle)
        }

        foreach ($tableStyle in $TableStyles.Values)
        {
            $documentTableStyle = Get-WordTableStyle -TableStyle $tableStyle -XmlDocument $XmlDocument
            [ref] $null = $documentStyles.AppendChild($documentTableStyle)
        }

        return $xmlDocument
    }
}

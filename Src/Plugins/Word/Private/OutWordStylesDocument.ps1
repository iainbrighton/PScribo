function OutWordStylesDocument
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
        $XmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $XmlDocument.AppendChild($XmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $documentStyles = $XmlDocument.AppendChild($XmlDocument.CreateElement('w', 'styles', $xmlnsMain))

        ## Create default style
        $defaultStyle = $documentStyles.AppendChild($XmlDocument.CreateElement('w', 'style', $xmlnsMain))
        [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'paragraph')
        [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1')
        [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'Normal')
        $defaultStyleName = $defaultStyle.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
        [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'Normal')
        [ref] $null = $defaultStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain))

        foreach ($Style in $Styles.Values)
        {
            $documentParagraphStyle = GetWordStyle -Style $Style -XmlDocument $XmlDocument -Type Paragraph
            [ref] $null = $documentStyles.AppendChild($documentParagraphStyle)
            $documentCharacterStyle = GetWordStyle -Style $Style -XmlDocument $XmlDocument -Type Character
            [ref] $null = $documentStyles.AppendChild($documentCharacterStyle)
        }

        foreach ($tableStyle in $TableStyles.Values)
        {
            $documentTableStyle = GetWordTableStyle -TableStyle $tableStyle -XmlDocument $XmlDocument
            [ref] $null = $documentStyles.AppendChild($documentTableStyle)
        }

        return $XmlDocument
    }
}

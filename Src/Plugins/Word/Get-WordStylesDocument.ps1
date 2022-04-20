function Get-WordStylesDocument
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
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $xmlDocument.AppendChild($XmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $documentStyles = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'styles', $xmlns))

        ## Create the document default style
        $defaultStyle = $Styles[$Document.DefaultStyle]
        $docDefaults = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'docDefaults', $xmlns))
        $rPrDefault = $docDefaults.AppendChild($xmlDocument.CreateElement('w', 'rPrDefault', $xmlns))
        $rPr = Get-WordStyleRunPr -Style $defaultStyle -XmlDocument $XmlDocument
        [ref] $null = $rPrDefault.AppendChild($rPr)
        $pPrDefault = $docDefaults.AppendChild($xmlDocument.CreateElement('w', 'pPrDefault', $xmlns))
        $pPr = $pPrDefault.AppendChild($xmlDocument.CreateElement('w', 'pPr', $xmlns))
        $spacing = $pPr.AppendChild($xmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('after', $xmlns, '0')

        ## Create default style (will inherit from the document default style)
        $documentStyle = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'style', $xmlns))
        [ref] $null = $documentStyle.SetAttribute('type', $xmlns, 'paragraph')
        [ref] $null = $documentStyle.SetAttribute('default', $xmlns, '1')
        [ref] $null = $documentStyle.SetAttribute('styleId', $xmlns, $defaultStyle.Id)
        $documentStyleName = $documentStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlns))
        [ref] $null = $documentStyleName.SetAttribute('val', $xmlns, $defaultStyle.Id)
        [ref] $null = $documentStyle.AppendChild($xmlDocument.CreateElement('w', 'qFormat', $xmlns))

        ## Create default character style (will inherit from the document default style)
        $documentCharacterStyleId = '{0}Char' -f $defaultStyle.Id
        $documentCharacterStyle = $documentStyles.AppendChild($xmlDocument.CreateElement('w', 'style', $xmlns))
        [ref] $null = $documentCharacterStyle.SetAttribute('type', $xmlns, 'character')
        [ref] $null = $documentCharacterStyle.SetAttribute('default', $xmlns, '1')
        [ref] $null = $documentCharacterStyle.SetAttribute('styleId', $xmlns, 'name')
        $documentCharacterStyleName = $documentCharacterStyle.AppendChild($xmlDocument.CreateElement('w', 'name', $xmlns))
        [ref] $null = $documentCharacterStyleName.SetAttribute('val', $xmlns, $documentCharacterStyleId)

        $nonDefaultStyles = $Styles.Values | Where-Object { $_.Id -ne $defaultStyle.Id }
        foreach ($Style in $nonDefaultStyles)
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

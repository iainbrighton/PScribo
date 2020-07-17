function Get-WordHeaderFooterDocument
{
<#
    .SYNOPSIS
        Outputs Office Open XML footer document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $HeaderFooter,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Header')]
        [System.Management.Automation.SwitchParameter] $IsHeader,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Footer')]
        [System.Management.Automation.SwitchParameter] $IsFooter
    )
    process
    {
        ## Create the Style.xml document
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $headerFooterDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $headerFooterDocument.AppendChild($headerFooterDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))

        if ($IsHeader -eq $true)
        {
            $elementName = 'hdr'
        }
        elseif ($IsFooter -eq $true)
        {
            $elementName = 'ftr'
        }
        $element = $headerFooterDocument.CreateElement('w', $elementName, $xmlns)

        foreach ($subSection in $HeaderFooter.Sections)
        {
            switch ($subSection.Type)
            {
                'PScribo.Paragraph'
                {
                    [ref] $null = $element.AppendChild((Out-WordParagraph -Paragraph $subSection -XmlDocument $headerFooterDocument))
                }
                'PScribo.Table'
                {
                    Out-WordTable -Table $subSection -XmlDocument $headerFooterDocument -Element $element
                }
                'PScribo.BlankLine'
                {
                    Out-WordBlankLine -BlankLine $subSection -XmlDocument $headerFooterDocument -Element $element
                }
            }
        }

        [ref] $null = $headerFooterDocument.AppendChild($element)
        return $headerFooterDocument
    }
}

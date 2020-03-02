function Out-WordDocument
{
<#
    .SYNOPSIS
        Outputs Office Open XML document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param
    (
        # The PScribo document object to convert to a Word xml document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlnswpdrawing = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        $xmlnsdrawing = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        $xmlnspicture = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        $xmlnsrelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $xmlnsofficeword14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
        $xmlnsmath = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $documentXml = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'document', $xmlnsMain))
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:xml', 'http://www.w3.org/XML/1998/namespace')
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:pic', $xmlnspicture)
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:wp', $xmlnswpdrawing)
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:a', $xmlnsdrawing)
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:r', $xmlnsrelationships)
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:m', $xmlnsmath)
        [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:a14', $xmlnsofficeword14)
        $body = $documentXml.AppendChild($xmlDocument.CreateElement('w', 'body', $xmlnsMain))

        foreach ($subSection in $Document.Sections.GetEnumerator())
        {
            if ($subSection.Id.Length -gt 40)
            {
                $sectionId = '{0}[..]' -f $subSection.Id.Substring(0, 36)
            }
            else
            {
                $sectionId = $subSection.Id
            }
            $currentIndentationLevel = 1
            if ($null -ne $subSection.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $subSection.Level + 1
            }
            WriteLog -Message ($localized.PluginProcessingSection -f $subSection.Type, $sectionId) -Indent $currentIndentationLevel
            switch ($subSection.Type)
            {
                'PScribo.Section' {
                    Out-WordSection -Section $subSection -Element $body -XmlDocument $xmlDocument
                }
                'PScribo.Paragraph' {
                    [ref] $null = $body.AppendChild((Out-WordParagraph -Paragraph $subSection -XmlDocument $xmlDocument))
                }
                'PScribo.Image' {
                    $Images += @($subSection)
                    [ref] $null = $body.AppendChild((Out-WordImage -Image $subSection -XmlDocument $xmlDocument))
                }
                'PScribo.PageBreak' {
                    [ref] $null = $body.AppendChild((Out-WordPageBreak -PageBreak $subSection -XmlDocument $xmlDocument))
                }
                'PScribo.LineBreak' {
                    [ref] $null = $body.AppendChild((Out-WordLineBreak -LineBreak $subSection -XmlDocument $xmlDocument))
                }
                'PScribo.Table' {
                    Out-WordTable -Table $subSection -XmlDocument $xmlDocument -Element $body
                }
                'PScribo.TOC' {
                    [ref] $null = $body.AppendChild((Out-WordTOC -TOC $subSection -XmlDocument $xmlDocument))
                }
                'PScribo.BlankLine' {
                    Out-WordBlankLine -BlankLine $subSection -XmlDocument $xmlDocument -Element $body
                }
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        ## Last section's properties are a child element of body element
        $lastSectionOrientation = $Document.Sections |
            Where-Object { $_.Type -in 'PScribo.Section','PScribo.Paragraph' } |
                Select-Object -Last 1 -ExpandProperty Orientation
        if ($null -eq $lastSectionOrientation)
        {
            $lastSectionOrientation = $Document.Options['PageOrientation']
        }

        $sectionPrParams = @{
            PageHeight       = $Document.Options['PageHeight']
            PageWidth        = $Document.Options['PageWidth']
            PageMarginTop    = $Document.Options['MarginTop'];
            PageMarginBottom = $Document.Options['MarginBottom'];
            PageMarginLeft   = $Document.Options['MarginLeft'];
            PageMarginRight  = $Document.Options['MarginRight'];
            Orientation      = $lastSectionOrientation
        }
        [ref] $null = $body.AppendChild((Get-WordSectionPr @sectionPrParams -XmlDocument $xmlDocument))

        return $xmlDocument
    }
}

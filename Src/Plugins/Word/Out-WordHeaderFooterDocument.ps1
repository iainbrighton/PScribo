function Out-WordHeaderFooterDocument
{
    [CmdletBinding()]
    param
    (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        [Parameter(Mandatory)]
        [System.IO.Packaging.Package] $Package
    )
    process
    {
        ## Create headers
        if ($Document.Header.HasFirstPageHeader)
        {
            $firstPageHeaderUri = New-Object -TypeName System.Uri -ArgumentList ('/word/firstPageHeader.xml', [System.UriKind]::Relative)
            $firstPageHeaderXml = Get-WordHeaderFooterDocument -HeaderFooter $Document.Header.FirstPageHeader -IsHeader
            Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $firstPageHeaderUri)
            $headerPart = $Package.CreatePart($firstPageHeaderUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
            $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($headerPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
            $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
            Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $firstPageHeaderUri)
            $firstPageHeaderXml.Save($xmlWriter)
            $xmlWriter.Dispose()
            $streamWriter.Close()
        }

        if ($Document.Header.HasDefaultHeader)
        {
            $defaultHeaderUri = New-Object -TypeName System.Uri -ArgumentList ('/word/defaultHeader.xml', [System.UriKind]::Relative)
            $defaultHeaderXml = Get-WordHeaderFooterDocument -HeaderFooter $Document.Header.DefaultHeader -IsHeader
            Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $defaultHeaderUri)
            $headerPart = $Package.CreatePart($defaultHeaderUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')
            $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($headerPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
            $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
            Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $defaultHeaderUri)
            $defaultHeaderXml.Save($xmlWriter)
            $xmlWriter.Dispose()
            $streamWriter.Close()
        }

        ## Create footers
        if ($Document.Footer.HasFirstPageFooter)
        {
            $firstPageFooterUri = New-Object -TypeName System.Uri -ArgumentList ('/word/firstPageFooter.xml', [System.UriKind]::Relative)
            $firstPageFooterXml = Get-WordHeaderFooterDocument -HeaderFooter $Document.Footer.FirstPageFooter -IsFooter
            Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $firstPageFooterUri)
            $footerPart = $Package.CreatePart($firstPageFooterUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
            $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($footerPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
            $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
            Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $firstPageFooterUri)
            $firstPageFooterXml.Save($xmlWriter)
            $xmlWriter.Dispose()
            $streamWriter.Close()
        }

        if ($Document.Footer.HasDefaultFooter)
        {
            $defaultFooterUri = New-Object -TypeName System.Uri -ArgumentList ('/word/defaultFooter.xml', [System.UriKind]::Relative)
            $defaultFooterXml = Get-WordHeaderFooterDocument -HeaderFooter $Document.Footer.DefaultFooter -IsFooter
            Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $defaultFooterUri)
            $footerPart = $Package.CreatePart($defaultFooterUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
            $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($footerPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
            $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
            Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $defaultFooterUri)
            $defaultFooterXml.Save($xmlWriter)
            $xmlWriter.Dispose()
            $streamWriter.Close()
        }
    }
}

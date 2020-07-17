function Get-WordSectionPr
{
<#
    .SYNOPSIS
        Outputs Office Open XML section element to set page size and margins.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory)]
        [System.Single] $PageWidth,

        [Parameter(Mandatory)]
        [System.Single] $PageHeight,

        [Parameter(Mandatory)]
        [System.Single] $PageMarginTop,

        [Parameter(Mandatory)]
        [System.Single] $PageMarginLeft,

        [Parameter(Mandatory)]
        [System.Single] $PageMarginBottom,

        [Parameter(Mandatory)]
        [System.Single] $PageMarginRight,

        [Parameter(Mandatory)]
        [System.String] $Orientation,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        if ($Orientation -eq 'Portrait')
        {
            $alignedPageHeight = $PageHeight
            $alignedPageWidth = $PageWidth
        }
        elseif ($Orientation -eq 'Landscape')
        {
            $alignedPageHeight = $PageWidth
            $alignedPageWidth = $PageHeight
        }

        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlns)

        $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz', $xmlns))
        [ref] $null = $pgSz.SetAttribute('w', $xmlns, (ConvertTo-Twips -Millimeter $alignedPageWidth))
        [ref] $null = $pgSz.SetAttribute('h', $xmlns, (ConvertTo-Twips -Millimeter $alignedPageHeight))
        [ref] $null = $pgSz.SetAttribute('orient', $xmlns, $Orientation.ToLower())

        $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMar', $xmlns))
        [ref] $null = $pgMar.SetAttribute('top', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginTop))
        [ref] $null = $pgMar.SetAttribute('header', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginTop))
        [ref] $null = $pgMar.SetAttribute('bottom', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginBottom))
        [ref] $null = $pgMar.SetAttribute('footer', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginBottom))
        [ref] $null = $pgMar.SetAttribute('left', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginLeft))
        [ref] $null = $pgMar.SetAttribute('right', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginRight))

        if ($Document.Header.HasFirstPageHeader)
        {
            $headerReference = $sectPr.AppendChild($xmlDocument.CreateElement('w', 'headerReference', $xmlns))
            [ref] $null = $headerReference.SetAttribute('type', $xmlns, 'first')
            [ref] $null = $headerReference.SetAttribute('id', $xmlnsrelationships, 'rId3')
        }

        if ($Document.Header.HasDefaultHeader)
        {
            $headerReference = $sectPr.AppendChild($xmlDocument.CreateElement('w', 'headerReference', $xmlns))
            [ref] $null = $headerReference.SetAttribute('type', $xmlns, 'default')
            [ref] $null = $headerReference.SetAttribute('id', $xmlnsrelationships, 'rId4')
        }

        if ($Document.Footer.HasFirstPageFooter)
        {
            $footerReference = $sectPr.AppendChild($xmlDocument.CreateElement('w', 'footerReference', $xmlns))
            [ref] $null = $footerReference.SetAttribute('type', $xmlns, 'first')
            [ref] $null = $footerReference.SetAttribute('id', $xmlnsrelationships, 'rId5')
        }

        if ($Document.Footer.HasDefaultFooter)
        {
            $footerReference = $sectPr.AppendChild($xmlDocument.CreateElement('w', 'footerReference', $xmlns))
            [ref] $null = $footerReference.SetAttribute('type', $xmlns, 'default')
            [ref] $null = $footerReference.SetAttribute('id', $xmlnsrelationships, 'rId6')
        }

        if (-not $script:pscriboIsFirstSection)
        {
            [ref] $null = $sectPr.AppendChild($xmlDocument.CreateElement('w', 'titlePg', $xmlns))
            $script:pscriboIsFirstSection = $true
        }

        return $sectPr
    }
}

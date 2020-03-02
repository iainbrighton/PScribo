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
        [ref] $null = $pgMar.SetAttribute('bottom', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginBottom))
        [ref] $null = $pgMar.SetAttribute('left', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginLeft))
        [ref] $null = $pgMar.SetAttribute('right', $xmlns, (ConvertTo-Twips -Millimeter $PageMarginRight))

        return $sectPr
    }
}

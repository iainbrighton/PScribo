function GetWordSectionPr
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
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlnsMain);
        $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz', $xmlnsMain));
        [ref] $null = $pgSz.SetAttribute('w', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageWidth));
        [ref] $null = $pgSz.SetAttribute('h', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageHeight));
        [ref] $null = $pgSz.SetAttribute('orient', $xmlnsMain, $Orientation.ToLower());
        $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMar', $xmlnsMain));
        [ref] $null = $pgMar.SetAttribute('top', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageMarginTop));
        [ref] $null = $pgMar.SetAttribute('bottom', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageMarginBottom));
        [ref] $null = $pgMar.SetAttribute('left', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageMarginLeft));
        [ref] $null = $pgMar.SetAttribute('right', $xmlnsMain, (ConvertTo-Twips -Millimeter $PageMarginRight));
        return $sectPr;
    }
}

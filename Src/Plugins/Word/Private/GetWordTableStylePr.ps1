function GetWordTableStylePr
{
<#
    .SYNOPSIS
        Generates Word table style (tblStylePr) formatting properties for specified table style type
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory)]
        [System.Management.Automation.PSObject] $Style,

        [Parameter(Mandatory)]
        [ValidateSet('Header', 'Row', 'AlternateRow')]
        [System.String] $Type,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlnsMain)
        [ref] $null = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))

        switch ($Type)
        {
            'Header' {
                $tblStylePrType = 'firstRow'
            }
            'Row' {
                $tblStylePrType = 'band2Horz'
            }
            'AlternateRow' {
                $tblStylePrType = 'band1Horz'
            }
        }

        [ref] $null = $tblStylePr.SetAttribute('type', $xmlnsMain, $tblStylePrType)
        [ref] $null = $tblStylePr.AppendChild((GetWordStyleParagraphPr -Style $Style -XmlDocument $XmlDocument))
        [ref] $null = $tblStylePr.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument))
        [ref] $null = $tblStylePr.AppendChild((GetWordTableStyleCellPr -Style $Style -XmlDocument $XmlDocument))
        return $tblStylePr
    }
}

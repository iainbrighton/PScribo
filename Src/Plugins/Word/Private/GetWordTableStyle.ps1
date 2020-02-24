function GetWordTableStyle
{
<#
    .SYNOPSIS
        Generates Word Xml table style element from a PScribo document table style.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        ## PScribo document style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $tableStyle,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $Style = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
        [ref] $null = $Style.SetAttribute('type', $xmlnsMain, 'table')
        [ref] $null = $Style.SetAttribute('styleId', $xmlnsMain, $tableStyle.Id)
        $name = $Style.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
        [ref] $null = $name.SetAttribute('val', $xmlnsMain, $tableStyle.Id)
        $tblPr = $Style.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))
        $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyleRowBandSize', $xmlnsMain))
        [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlnsMain, 1)

        if ($tableStyle.BorderWidth -gt 0)
        {
            $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain))
            foreach ($border in @('top', 'bottom', 'start', 'end', 'insideH', 'insideV'))
            {
                $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain))
                [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertTo-Octips $tableStyle.BorderWidth))
                [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single')
                [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor))
            }
        }

        [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.HeaderStyle] -Type Header -XmlDocument $XmlDocument))
        [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.RowStyle] -Type Row -XmlDocument $XmlDocument))
        [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.AlternateRowStyle] -Type AlternateRow -XmlDocument $XmlDocument))
        return $Style
    }
}

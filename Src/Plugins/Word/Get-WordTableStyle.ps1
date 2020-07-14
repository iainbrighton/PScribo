function Get-WordTableStyle
{
<#
    .SYNOPSIS
        Generates Word Xml table style element from a PScribo document table style.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        ## PScribo document table style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TableStyle,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $style = $XmlDocument.CreateElement('w', 'style', $xmlns)
        [ref] $null = $style.SetAttribute('type', $xmlns, 'table')
        [ref] $null = $style.SetAttribute('customStyle', $xmlns, '1')
        [ref] $null = $style.SetAttribute('styleId', $xmlns, $tableStyle.Id)

        $name = $style.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlns))
        [ref] $null = $name.SetAttribute('val', $xmlns, $TableStyle.Id)

        $basedOn = $style.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlns))
        [ref] $null = $basedOn.SetAttribute('val', $xmlns, 'TableNormal')

        $tblPr = $style.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlns))
        $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyleRowBandSize', $xmlns))
        [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlns, 1)

        if ($TableStyle.BorderWidth -gt 0)
        {
            $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlns))
            foreach ($border in @('top', 'bottom', 'start', 'end', 'insideH', 'insideV'))
            {
                $borderType = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlns))
                [ref] $null = $borderType.SetAttribute('sz', $xmlns, (ConvertTo-Octips $tableStyle.BorderWidth))
                [ref] $null = $borderType.SetAttribute('val', $xmlns, 'single')
                $borderColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $tableStyle.BorderColor)
                [ref] $null = $borderType.SetAttribute('color', $xmlns, (ConvertTo-WordColor -Color $borderColor))
            }
        }

        [ref] $null = $Style.AppendChild((Get-WordTableStylePr -Style $Document.Styles[$TableStyle.HeaderStyle] -Type HeaderRow -XmlDocument $XmlDocument))
        [ref] $null = $Style.AppendChild((Get-WordTableStylePr -Style $Document.Styles[$TableStyle.HeaderStyle] -Type HeaderColumn -XmlDocument $XmlDocument))
        [ref] $null = $Style.AppendChild((Get-WordTableStylePr -Style $Document.Styles[$TableStyle.RowStyle] -Type Row -XmlDocument $XmlDocument))
        [ref] $null = $Style.AppendChild((Get-WordTableStylePr -Style $Document.Styles[$TableStyle.AlternateRowStyle] -Type AlternateRow -XmlDocument $XmlDocument))

        return $style
    }
}

function Get-WordTableStylePr
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
        [ValidateSet('HeaderRow', 'HeaderColumn', 'Row', 'AlternateRow')]
        [System.String] $Type,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        if ($Type -eq 'HeaderRow')
        {
            $tblStylePrType = 'firstRow'
        }
        elseif ($Type -eq 'HeaderColumn')
        {
            $tblStylePrType = 'firstCol'
        }
        elseif ($Type -eq 'Row')
        {
            $tblStylePrType = 'band1Horz'
        }
        elseif ($Type -eq 'AlternateRow')
        {
            $tblStylePrType = 'band2Horz'
        }

        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlns)
        [ref] $null = $tblStylePr.SetAttribute('type', $xmlns, $tblStylePrType)

        $rPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlns))
        $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlns))
        [ref] $null = $rFonts.SetAttribute('ascii', $xmlns, $Style.Font[0])
        [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlns, $Style.Font[0])

        if ($Style.Bold)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlns))
        }
        if ($Style.Underline)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlns))
        }
        if ($Style.Italic)
        {
            [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlns))
        }

        $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlns))
        [ref] $null = $color.SetAttribute('val', $xmlns, (ConvertTo-WordColor -Color $Style.Color))

        $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlns))
        [ref] $null = $sz.SetAttribute('val', $xmlns, ($Style.Size * 2))

        $tblPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlns))
        if (-not [System.String]::IsNullOrEmpty($Style.BackgroundColor))
        {
            $tcPr = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlns))
            $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlns))
            [ref] $null = $shd.SetAttribute('val', $xmlns, 'clear')
            [ref] $null = $shd.SetAttribute('color', $xmlns, 'auto')
            $backgroundColor = ConvertTo-WordColor -Color $Style.BackgroundColor
            [ref] $null = $shd.SetAttribute('fill', $xmlns, $backgroundColor)
        }

        return $tblStylePr
    }
}

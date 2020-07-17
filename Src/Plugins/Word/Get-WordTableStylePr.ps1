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

        $rPr = Get-WordStyleRunPr -Style $Style -XmlDocument $XmlDocument
        [ref] $null = $tblStylePr.AppendChild($rPr)

        $null = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlns))
        $tcPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlns))
        if (-not [System.String]::IsNullOrEmpty($Style.BackgroundColor))
        {
            $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlns))
            [ref] $null = $shd.SetAttribute('val', $xmlns, 'clear')
            [ref] $null = $shd.SetAttribute('color', $xmlns, 'auto')
            $backgroundColor = ConvertTo-WordColor -Color $Style.BackgroundColor
            [ref] $null = $shd.SetAttribute('fill', $xmlns, $backgroundColor)
        }

        $pPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlns))
        if ($Style.Align.ToLower() -eq 'justify')
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, 'distribute')
        }
        else
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, $Style.Align.ToLower())
        }

        return $tblStylePr
    }
}

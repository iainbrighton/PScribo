function Get-WordTablePr
{
<#
    .SYNOPSIS
    Creates a scaffold Word <w:tblPr> element
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tableStyle = $Document.TableStyles[$Table.Style]
        $tblPr = $XmlDocument.CreateElement('w', 'tblPr', $xmlns)

        $tblStyle = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyle', $xmlns))
        [ref] $null = $tblStyle.SetAttribute('val', $xmlns, $tableStyle.Id)

        $tableWidthRenderPct = $Table.Width
        if ($Table.Tabs -gt 0)
        {
            $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlns))
            [ref] $null = $tblInd.SetAttribute('w', $xmlns, (720 * $Table.Tabs))

            ## We now need to deal with tables being pushed outside the page margin
            $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options['PageMarginLeft'] + $Document.Options['PageMarginRight'])
            $indentWidthMm = ConvertTo-Mm -Point ($Table.Tabs * 36)
            $tableRenderMm = (($pageWidthMm / 100) * $Table.Width) + $indentWidthMm
            if ($tableRenderMm -gt $pageWidthMm)
            {
                ## We've over-flowed so need to work out the maximum percentage
                $maxTableWidthMm = $pageWidthMm - $indentWidthMm
                $tableWidthRenderPct = [System.Math]::Round(($maxTableWidthMm / $pageWidthMm) * 100, 2)
                WriteLog -Message ($localized.TableWidthOverflowWarning -f $tableWidthRenderPct) -IsWarning
            }
        }

        if ($Table.ColumnWidths -or ($Table.Width -gt 0 -and $Table.Width -lt 100))
        {
            $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlns))
            [ref] $null = $tblLayout.SetAttribute('type', $xmlns, 'fixed')
        }

        $tableAlignment = @{ Left = 'start'; Center = 'center'; Right = 'end'; }
        $jc =  $tblPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlns))
        [ref] $null = $jc.SetAttribute('val', $xmlns, $tableAlignment[$tableStyle.Align])

        if ($Table.Width -gt 0)
        {
            $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlns))
            [ref] $null = $tblW.SetAttribute('w', $xmlns, ($tableWidthRenderPct * 50))
            [ref] $null = $tblW.SetAttribute('type', $xmlns, 'pct')
        }

        $tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlns))
        $isFirstRow = ($Table.IsKeyedList -eq $true -or $Table.IsList -eq $false) -as [System.Int32]
        $isFirstColumn = ($Table.IsKeyedList -eq $true -or $Table.IsList -eq $true) -as [System.Int32]
        [ref] $null = $tblLook.SetAttribute('firstRow', $xmlns, $isFirstRow)
        [ref] $null = $tblLook.SetAttribute('lastRow', $xmlns, 0)
        [ref] $null = $tblLook.SetAttribute('firstColumn', $xmlns, $isFirstColumn)
        [ref] $null = $tblLook.SetAttribute('lastColumn', $xmlns, 0)
        [ref] $null = $tblLook.SetAttribute('noHBand', $xmlns, 0)
        [ref] $null = $tblLook.SetAttribute('noVBand', $xmlns, 1)

        if ($tableStyle.BorderWidth -gt 0)
        {
            $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlns))
            foreach ($border in @('top', 'bottom', 'start', 'end', 'insideH', 'insideV'))
            {
                $borderType = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlns))
                [ref] $null = $borderType.SetAttribute('sz', $xmlns, (ConvertTo-Octips $tableStyle.BorderWidth))
                [ref] $null = $borderType.SetAttribute('val', $xmlns, 'single')
                [ref] $null = $borderType.SetAttribute('color', $xmlns, (ConvertTo-WordColor -Color $tableStyle.BorderColor))
            }
        }

        $tblCellMar = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblCellMar', $xmlns))
        $top = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'top', $xmlns))
        [ref] $null = $top.SetAttribute('w', $xmlns, (ConvertTo-Twips $tableStyle.PaddingTop))
        [ref] $null = $top.SetAttribute('type', $xmlns, 'dxa')
        $start = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlns))
        [ref] $null = $start.SetAttribute('w', $xmlns, (ConvertTo-Twips $tableStyle.PaddingLeft))
        [ref] $null = $start.SetAttribute('type', $xmlns, 'dxa')
        $bottom = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlns))
        [ref] $null = $bottom.SetAttribute('w', $xmlns, (ConvertTo-Twips $tableStyle.PaddingBottom))
        [ref] $null = $bottom.SetAttribute('type', $xmlns, 'dxa')
        $end = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'end', $xmlns))
        [ref] $null = $end.SetAttribute('w', $xmlns, (ConvertTo-Twips $tableStyle.PaddingRight))
        [ref] $null = $end.SetAttribute('type', $xmlns, 'dxa')

        $spacing = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('before', $xmlns, 72)
        [ref] $null = $spacing.SetAttribute('after', $xmlns, 72)

        return $tblPr
    }
}

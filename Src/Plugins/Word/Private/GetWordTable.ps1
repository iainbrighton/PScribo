function GetWordTable
{
<#
    .SYNOPSIS
    Creates a scaffold Word <w:tbl> element
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
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tableStyle = $Document.TableStyles[$Table.Style]
        $tbl = $XmlDocument.CreateElement('w', 'tbl', $xmlnsMain)
        $tblPr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))

        if ($Table.Tabs -gt 0)
        {
            $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlnsMain))
            [ref] $null = $tblInd.SetAttribute('w', $xmlnsMain, (720 * $Table.Tabs))
        }

        if ($Table.ColumnWidths)
        {
            $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain))
            [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'fixed')
        }
        elseif ($Table.Width -eq 0)
        {
            $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain))
            [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'autofit')
        }

        if ($Table.Width -gt 0)
        {
            $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlnsMain))
            [ref] $null = $tblW.SetAttribute('type', $xmlnsMain, 'pct')
            $tableWidthRenderPct = $Table.Width

            if ($Table.Tabs -gt 0)
            {
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
            [ref] $null = $tblW.SetAttribute('w', $xmlnsMain, $tableWidthRenderPct * 50)
        }

        $spacing = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 72)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 72)

        #$tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlnsMain));
        #[ref] $null = $tblLook.SetAttribute('val', $xmlnsMain, '04A0');
        #[ref] $null = $tblLook.SetAttribute('firstRow', $xmlnsMain, 1);
        ## <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
        #$tblStyle = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyle', $xmlnsMain));
        #[ref] $null = $tblStyle.SetAttribute('val', $xmlnsMain, $Table.Style);

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

        $tblCellMar = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblCellMar', $xmlnsMain))
        $top = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'top', $xmlnsMain))
        [ref] $null = $top.SetAttribute('w', $xmlnsMain, (ConvertTo-Twips $tableStyle.PaddingTop))
        [ref] $null = $top.SetAttribute('type', $xmlnsMain, 'dxa')
        $left = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlnsMain))
        [ref] $null = $left.SetAttribute('w', $xmlnsMain, (ConvertTo-Twips $tableStyle.PaddingLeft))
        [ref] $null = $left.SetAttribute('type', $xmlnsMain, 'dxa')
        $bottom = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain))
        [ref] $null = $bottom.SetAttribute('w', $xmlnsMain, (ConvertTo-Twips $tableStyle.PaddingBottom))
        [ref] $null = $bottom.SetAttribute('type', $xmlnsMain, 'dxa')
        $right = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'end', $xmlnsMain))
        [ref] $null = $right.SetAttribute('w', $xmlnsMain, (ConvertTo-Twips $tableStyle.PaddingRight))
        [ref] $null = $right.SetAttribute('type', $xmlnsMain, 'dxa')

        $tblGrid = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblGrid', $xmlnsMain))
        for ($i = 0; $i -lt $Table.Columns.Count; $i++)
        {
            [ref] $null = $tblGrid.AppendChild($XmlDocument.CreateElement('w', 'gridCol', $xmlnsMain))
        }

        return $tbl
    }
}

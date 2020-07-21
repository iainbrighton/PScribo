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

        $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlns))
        [ref] $null = $tblInd.SetAttribute('w', $xmlns, (720 * $Table.Tabs))
        [ref] $null = $tblInd.SetAttribute('type', $xmlns, 'dxa')

        $getWordTableRenderWidthMmParams = @{
            TableWidth  = $Table.Width
            Tabs        = $Table.Tabs
            Orientation = $Table.Orientation
        }
        $tableRenderWidthMm = Get-WordTableRenderWidthMm @getWordTableRenderWidthMmParams
        $tableRenderWidthTwips = ConvertTo-Twips -Millimeter $tableRenderWidthMm

        $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlns))
        if ($Table.ColumnWidths -or ($Table.Width -gt 0 -and $Table.Width -le 100))
        {
            [ref] $null = $tblLayout.SetAttribute('type', $xmlns, 'fixed')
        }
        else
        {
            [ref] $null = $tblLayout.SetAttribute('type', $xmlns, 'autofit')
        }

        $tableAlignment = @{ Left = 'start'; Center = 'center'; Right = 'end'; }
        $jc =  $tblPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlns))
        [ref] $null = $jc.SetAttribute('val', $xmlns, $tableAlignment[$tableStyle.Align])

        if ($Table.Width -gt 0)
        {
            $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlns))
            [ref] $null = $tblW.SetAttribute('w', $xmlns, $tableRenderWidthTwips)
            [ref] $null = $tblW.SetAttribute('type', $xmlns, 'dxa')
        }

        $tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlns))
        $isFirstRow = ($Table.IsKeyedList -eq $true -or $Table.IsList -eq $false)
        $isFirstColumn = ($Table.IsKeyedList -eq $true -or $Table.IsList -eq $true)
        ## LibreOffice requires legacy conditional formatting value (#99)
        $val = Get-WordTableConditionalFormattingValue -HasFirstRow:$isFirstRow -HasFirstColumn:$isFirstColumn -NoVerticalBand
        [ref] $null = $tblLook.SetAttribute('val', $xmlns, ('{0:x4}' -f $val))
        [ref] $null = $tblLook.SetAttribute('firstRow', $xmlns, ($isFirstRow -as [System.Int32]))
        [ref] $null = $tblLook.SetAttribute('lastRow', $xmlns, '0')
        [ref] $null = $tblLook.SetAttribute('firstColumn', $xmlns, ($isFirstColumn -as [System.Int32]))
        [ref] $null = $tblLook.SetAttribute('lastColumn', $xmlns, '0')
        [ref] $null = $tblLook.SetAttribute('noHBand', $xmlns, '0')
        [ref] $null = $tblLook.SetAttribute('noVBand', $xmlns, '1')

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

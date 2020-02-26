function ConvertTo-PScriboFormattedKeyedListTable
{
<#
    .SYNOPSIS
        Creates a formatted keyed list table (a key'd column per object) for plugin output/rendering.

    .NOTES
        Maintains backwards compatibility with other plugins that do not require styling/formatting.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject[]])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Table
    )
    begin
    {
        $hasColumnWidths = ($null -ne $Table.ColumnWidths)
        $objectKey = $Table.ListKey
    }
    process
    {
        $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderRow -HasHeaderColumn

        ## Output the header cells
        $headerRow = New-PScriboFormattedTableRow -TableStyle $Table.Style -IsHeaderRow

        if ($Table.DisplayListKey)
        {
            $newPScriboFormattedTableRowHeaderCellParams = @{
                Content = $Table.ListKey
            }
        }
        else
        {
            $newPScriboFormattedTableRowHeaderCellParams = @{
                Content = ' '
            }
        }
        $listKeyHeaderCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
        $null = $headerRow.Cells.Add($listKeyHeaderCell)

        ## Add all the object key values
        for ($o = 0; $o -lt $Table.Rows.Count; $o++)
        {
            $newPScriboFormattedTableRowHeaderCellParams = @{
                Content = $Table.Rows[$o].PSObject.Properties[$objectKey].Value
            }
            $objectKeyCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
            $null = $headerRow.Cells.Add($objectKeyCell)
        }
        $null = $formattedTable.Rows.Add($headerRow)

        $isAlternateRow = $false
        ## Output remaining object properties (one property per row)
        foreach ($column in $Table.Columns)
        {
            if ((-not $column.EndsWith('__Style', 'CurrentCultureIgnoreCase')) -and
                ($column -ne $objectKey))
            {
                ## Add the object property column
                $newPScriboFormattedTableRowParams = @{
                    TableStyle = $Table.Style;
                    IsAlternateRow = $isAlternateRow
                }
                $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

                ## Output the column header cell (property name) as header style
                $newPScriboFormattedTableColumnCellParams = @{
                    Content = $column
                }
                if ($hasColumnWidths) {
                    $newPScriboFormattedTableColumnCellParams['Width'] = $Table.ColumnWidths[0]
                }
                $columnCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableColumnCellParams
                $null = $row.Cells.Add($columnCell)

                ## Add the property value for all other objects
                for ($o = 0; $o -lt $Table.Rows.Count; $o++)
                {
                    $propertyValue = $Table.Rows[$o].PSObject.Properties[$column].Value
                    $newPScriboFormattedTableRowValueCellParams = @{
                        Content = $propertyValue
                    }
                    if ([System.String]::IsNullOrEmpty($propertyValue)) {
                        $newPScriboFormattedTableRowValueCellParams['Content'] = $null
                    }
                    $propertyStyleName = '{0}__Style' -f $column
                    $hasStyleProperty = $Table.Rows[$o].PSObject.Properties.Name.Contains($propertyStyleName)
                    if ($hasStyleProperty) {
                        $newPScriboFormattedTableRowValueCellParams['Style'] = $Table.Rows[$o].PSObject.Properties[$propertyStyleName].Value
                    }
                    if ($hasColumnWidths) {
                        $newPScriboFormattedTableRowValueCellParams['Width'] = $Table.ColumnWidths[$o+1]
                    }
                    $valueCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowValueCellParams
                    $null = $row.Cells.Add($valueCell)
                }

                $null = $formattedTable.Rows.Add($row)
                $isAlternateRow = -not $isAlternateRow
            }
        }
        Write-Output -InputObject $formattedTable
    }
}

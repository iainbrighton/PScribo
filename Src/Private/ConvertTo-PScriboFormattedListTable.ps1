function ConvertTo-PScriboFormattedListTable
{
<#
    .SYNOPSIS
        Creates a formatted list table for plugin output/rendering.

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
    }
    process
    {
        for ($r = 0; $r -lt $Table.Rows.Count; $r++)
        {
            ## We have one table per object
            $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderColumn
            $objectProperties = $Table.Rows[$r].PSObject.Properties

            for ($c = 0; $c -lt $Table.Columns.Count; $c++)
            {
                $column = $Table.Columns[$c]
                $newPScriboFormattedTableRowParams = @{
                    TableStyle = $Table.Style;
                    Style = $Table.Rows[$r].'__Style'
                    IsAlternateRow = ($c % 2) -ne 0
                }
                $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

                ## Add each property name (as header style)
                $newPScriboFormattedTableRowHeaderCellParams = @{
                    Content = $column
                }
                if ($hasColumnWidths)
                {
                    $newPScriboFormattedTableRowHeaderCellParams['Width'] = $Table.ColumnWidths[0]
                }
                $headerCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
                $null = $row.Cells.Add($headerCell)

                ## Add each property value
                $propertyValue = $objectProperties[$column].Value
                if ([System.String]::IsNullOrEmpty($propertyValue))
                {
                    $propertyValue = $null
                }

                $newPScriboFormattedTableRowValueCellParams = @{
                    Content = $propertyValue
                }

                $propertyStyleName = '{0}__Style' -f $column
                $hasStyleProperty = $objectProperties.Name.Contains($propertyStyleName)
                if ($hasStyleProperty)
                {
                    $newPScriboFormattedTableRowValueCellParams['Style'] = $objectProperties[$propertyStyleName].Value
                }
                if ($hasColumnWidths)
                {
                    $newPScriboFormattedTableRowValueCellParams['Width'] = $Table.ColumnWidths[1]
                }
                $valueCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowValueCellParams
                $null = $row.Cells.Add($valueCell)

                $null = $formattedTable.Rows.Add($row)
            }
            Write-Output -InputObject $formattedTable
        }
    }
}

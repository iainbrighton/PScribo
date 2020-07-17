function ConvertTo-PScriboFormattedTable
{
<#
    .SYNOPSIS
        Creates a formatted standard table for plugin output/rendering.

    .NOTES
        Maintains backwards compatibility with other plugins that do not require styling/formatting.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject])]
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
        $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderRow

        # Output the header row and header cells
        $headerRow = New-PScriboFormattedTableRow -TableStyle $Table.Style -IsHeaderRow
        for ($h = 0; $h -lt $Table.Columns.Count; $h++)
        {
            $newPScriboFormattedTableHeaderCellParams = @{
                Content    = $Table.Columns[$h]
            }
            if ($hasColumnWidths)
            {
                $newPScriboFormattedTableHeaderCellParams['Width'] = $Table.ColumnWidths[$h]
            }
            $cell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableHeaderCellParams
            $null = $headerRow.Cells.Add($cell)
        }
        $null = $formattedTable.Rows.Add($headerRow)

        ## Output each object row
        for ($r = 0; $r -lt $Table.Rows.Count; $r++)
        {
            $objectProperties = $Table.Rows[$r].PSObject.Properties

            $newPScriboFormattedTableRowParams = @{
                TableStyle = $Table.Style;
                Style = $Table.Rows[$r].'__Style'
                IsAlternateRow = ($r % 2 -ne 0 )
            }
            $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

            ## Output object row's cells
            for ($c = 0; $c -lt $Table.Columns.Count; $c++)
            {
                $propertyName = $Table.Columns[$c]
                $propertyStyleName = '{0}__Style' -f $propertyName;
                $hasStyleProperty = $objectProperties.Name.Contains($propertyStyleName)

                $propertyValue = $objectProperties[$propertyName].Value

                $newPScriboFormattedTableRowCellParams = @{
                    Content = $propertyValue
                }
                if ([System.String]::IsNullOrEmpty($propertyValue))
                {
                    $newPScriboFormattedTableRowCellParams['Content'] = $null
                }

                if ($hasColumnWidths)
                {
                    $newPScriboFormattedTableRowCellParams['Width'] = $Table.ColumnWidths[$c]
                }
                if ($hasStyleProperty)
                {
                    $newPScriboFormattedTableRowCellParams['Style'] = $objectProperties[$propertyStyleName].Value # | Where-Object Name -eq $propertyStyleName | Select-Object -ExpandProperty Value
                }

                $cell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowCellParams
                $null = $row.Cells.Add($cell)
            }
            $null = $formattedTable.Rows.Add($row)
        }
        Write-Output -InputObject $formattedTable
    }
}

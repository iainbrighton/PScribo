function OutHtmlTable
{
<#
    .SYNOPSIS
        Output formatted Html <table> from PScribo.Table object.
    .NOTES
        One table is output per table row with the -List parameter.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {

        [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        if ($Table.IsKeyedList)
        {
            $formattedTable = GetHtmlFormattedTable -Table $Table
            [ref] $null = $tableBuilder.Append($formattedTable)
        }
        elseif ($Table.IsList)
        {
            for ($r = 0; $r -lt $Table.Rows.Count; $r++)
            {
                $row = $Table.Rows[$r]
                $cloneTable = Copy-Object -InputObject $Table
                $cloneTable.Rows = @($row)
                $formattedTable = GetHtmlFormattedTable -Table $cloneTable
                [ref] $null = $tableBuilder.Append($formattedTable)
                ## Add a space between each table to mirror Word output rendering
                [ref] $null = $tableBuilder.AppendLine('<p />');
            } #end foreach row
        }
        else
        {
            $formattedTable = GetHtmlFormattedTable -Table $Table
            [ref] $null = $tableBuilder.Append($formattedTable)
        }
        return $tableBuilder.ToString()
    }
}

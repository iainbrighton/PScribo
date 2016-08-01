Import-Module PScribo -Force;

$example14 = Document -Name 'PScribo Example 14' {
    <#
        By default, tables are configured to 100% of available width and to autofit
        the table's cell contents. You can override the table width, by specifying the
        -Width parameter on the 'Table' cmdlet. The width is always set as a percentage
        of available space.

        The following example creates a table with its with set at 66% of the available
        space.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status -Headers 'Name','Display Name','State' -Width 66
}
$example14 | Export-Document -Format Html -Path ~\Desktop

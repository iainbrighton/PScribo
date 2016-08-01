Import-Module PScribo -Force;

$example13 = Document -Name 'PScribo Example 13' {
    <#
        To override the column headers, the -Columns and -Headers parameters can be used
        together.

        The -Headers parameter specifies the heading names to apply to the corresponding
        (positional) -Columns parameter. This permits you to include spaces in the headers
        and/or change the properties' display values.

        The following example overrides the column names with the values supplied in the
        -Headers parameter - a space is introduced in the 'DisplayName' property and the
        'Status' property is displayed as 'State' instead.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status -Headers Name,'Display Name',State
}
$example13 | Export-Document -Format Html -Path ~\Desktop

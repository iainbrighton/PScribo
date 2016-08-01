Import-Module PScribo -Force;

$example12 = Document -Name 'PScribo Example 12' {
    <#
        The 'Table' cmdlet supports filtering of object properties by using the -Columns
        parameter. The -Columns parameter specifies the object properties you wish to
        display in the table, in order required.

        NOTE: The -Columns parameter is required if you wish to override the values
        displayed in column headings.

        The following example displays the Name, DisplayName and Status properties of
        all services sent down the pipeline, ignoring all other properties.

        NOTE: The columns are displayed in the exact order that they are listed in
        the -Columns parameter.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status
}
$example12 | Export-Document -Format Html -Path ~\Desktop

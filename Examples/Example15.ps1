[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example15 = Document -Name 'PScribo Example 15' {

    <#
        You can optionally set the width of each column by using the -ColumnWidths
        parameter on the 'Table' cmdlet.

        NOTE: You can only specify column widths if you also specify the -Columns
              parameter.

        Just like the Table -Width parameter, the column widths are specified in
        percentages (of overall table width).

        NOTE: The total of all columns widths must total exactly 100 (%).

        The following example retrieves all local services, displaying the Name,
        DisplayName and Status properties. The column width for the Name property
        is set to 30%, the column width for the DisplayName property is set to 50%
        and the Status property column width set to 20%.

        NOTE: It is recommended that you set column widths on all tables to ensure
              consistent formatting between HTML and Word outputs.

        NOTE: Column widths are not implemented for Text output rendering as they
              are not supported/implemented in the 'Format-Table' cmdlet.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status -ColumnWidths 30,50,20
}
$example15 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

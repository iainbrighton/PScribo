[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example19 = Document -Name 'PScribo Example 19' {

    <#
        Column widths can also be specified on -List tables. The total of the
        column widths must still total 100 (%), but only 2 widths are supported.

        NOTE: -List view also supports the standard 'Table' -Width parameter.

        The following example creates a table for the first 5 services
        registered on the local machine, and sets the first (property) column
        to 40% and the second (value) column to 60%.
    #>
    Get-Service | Select-Object -First 5 | Table -List -ColumnWidths 40,60
}
$example19 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

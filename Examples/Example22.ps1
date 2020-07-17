[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example22 = Document -Name 'PScribo Example 22' {

    <#
        A grid can be applied to a table by setting the -BorderWidth property. The
        border width is specified in points (pt).

        NOTE: Table borders apply to both standard tables and -List view tables.

        Optionally, you may set a border/grid color with the -BorderColor property.
        If the border color is not specified, it will default to black (#000).

        The following example creates a table with an orange/red grid/border around
        all cells.
    #>
    TableStyle -Name 'BasicGrid' -HeaderStyle Normal -RowStyle Normal -BorderWidth 1 -BorderColor OrangeRed
    Get-Service | Select-Object -Property Name,DisplayName,Status -First 3 | Table -Style BasicGrid
}
$example22 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

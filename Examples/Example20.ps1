[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example20 = Document -Name 'PScribo Example 20' {

    <#
        Hashtables can also be used in -List view along with the -Width
        and -ColumnWidths parameters.

        The following example creates three tables (one per hashtable) in -List
        view,  from a collection of manually constructed hashtables
    #>
    $hashtableArray = @(
        [Ordered] @{ 'Column 1' = 'Some random text'; 'Column2' = 345; 'Custom Property' = $true; }
        [Ordered] @{ 'Column 1' = 'Random some text'; 'Column2' = 16; 'Custom Property' = $false; }
        [Ordered] @{ 'Column 1' = 'Text random some'; 'Column2' = 1GB; 'Custom Property' = $true; }
    )
    Table -Hashtable $hashtableArray -Width 50 -ColumnWidths 40,60 -List
}
$example20 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

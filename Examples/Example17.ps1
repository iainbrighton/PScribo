[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example17 = Document -Name 'PScribo Example 17' {

    <#
        To create a multi-row table from hashtables, simply create a collection/
        array of ordered hashtables.

        NOTE: You must ensure that the hashtable keys are the same on all hashtables
              in the collection/array as only the first object is enumerated.

        The following example creates an table with a three rows from an array of
        ordered hashtables.
    #>
    $hashtableArray = @(
        [Ordered] @{ 'Column 1' = 'Some random text'; 'Column2' = 345; 'Custom Property' = $true; }
        [Ordered] @{ 'Column 1' = 'Random some text'; 'Column2' = 16; 'Custom Property' = $false; }
        [Ordered] @{ 'Column 1' = 'Text random some'; 'Column2' = 1GB; 'Custom Property' = $true; }
    )
    Table -Hashtable $hashtableArray
}
$example17 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

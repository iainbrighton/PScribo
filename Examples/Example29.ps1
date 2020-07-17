[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example29 = Document -Name 'PScribo Example 29' {

    <#
        Custom styling can also be applied to individual keys of one or more
        hashtables.

        To apply a style to a particular hashtable key, just ensure that a
        '<KeyName>__Style' key is also defined with the style name you wish to to apply
        to the particular table cell.

        For example, to apply a "MyCustomStyle" to a 'Name' key, ensure that a
        'Name__Style' key exists in the hashtable with a value of "MyCustomStyle".

        In the following example, the the middle two rows where key 'Custom Property' is
        set to $false will be highlighted with the in the array of hashtables
        will be styled with the 'Warning' custom style.

        NOTE: How the 'Custom Property__Style' key is added to a hashtable is up to you!
    #>
    Style -Name Warning -Color White -BackgroundColor Firebrick

    $hashtableArray = @(
        [Ordered] @{ 'Column 1' = 'Some random text'; 'Column2' = 345; 'Custom Property' = $true; }
        [Ordered] @{ 'Column 1' = 'Random some text'; 'Column2' = 16; 'Custom Property' = $false; 'Custom Property__Style' = 'Warning'; }
        [Ordered] @{ 'Column 1' = 'Random text some'; 'Column2' = 241; 'Custom Property' = $false; 'Custom Property__Style' = 'Warning'; }
        [Ordered] @{ 'Column 1' = 'Text random some'; 'Column2' = 1GB; 'Custom Property' = $true; }
    )
    Table -Hashtable $hashtableArray
}
$example29 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

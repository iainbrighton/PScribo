[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example27 = Document -Name 'PScribo Example 27' {

    <#
        Custom styling can also be applied to rows of one or more hashtables.

        The following defines a custom style that we will apply to particular objects
        within the collection.
    #>
    Style -Name StoppedService -Color White -BackgroundColor Firebrick

    <#
        To apply a style to a particular hashtable, just ensure that a '__Style' key
        is defined with the required style name to apply to the row!

        In the following example, the first and last rows in the array of hashtables
        will be styled with the 'StoppedService' style.

        NOTE: How the '__Style' key is added to a hashtable is up to you. It could
              be added programatically etc.
    #>
    $hashtableArray = @(
        [Ordered] @{ 'Column 1' = 'Some random text'; 'Column2' = 345; 'Custom Property' = $true; '__Style' = 'StoppedService'; }
        [Ordered] @{ 'Column 1' = 'Random some text'; 'Column2' = 16; 'Custom Property' = $false; }
        [Ordered] @{ 'Column 1' = 'Random text some'; 'Column2' = 241; 'Custom Property' = $true; }
        [Ordered] @{ 'Column 1' = 'Text random some'; 'Column2' = 1GB; 'Custom Property' = $true; '__Style' = 'StoppedService'; }
    )
    Table -Hashtable $hashtableArray
}
$example27 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

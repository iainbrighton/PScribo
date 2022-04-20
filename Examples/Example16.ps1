[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example16 = Document -Name 'PScribo Example 16' {

    <#
        If you do not have a collection/array of objects that you wish to create
        a table from, PScribo supports creating a table from single ordered hashtable
        or an array of ordered hashtables.

        NOTE: You cannot pipe a single hashtable or an array of hashtables to the
              'Table' cmdlet. You MUST pass the hashtable(s) via the -HashTable parameter.
              If you pipe the hashtable(s), the hashtable object's properties are
              displayed and not the hashtable key/value pairs!

        Creating a hashtable does permit the utilisation of spaces in the key names,
        requiring that the -Headers parameter does not necessarily need to be used.

        NOTE: This is not a standard hashtable, but a System.Collections.Specialized.OrderedDictionary
              object. These can be created with the [Ordered] attribute on the hashtable
              declaration, e.g. [Ordered] @{ Key1 = Value1; Key2 = Value2; }

        The following example creates a table with a single row with the hashtable keys
        used as the header values and the corresponding values as the first table row.
    #>
    $hashtable = [Ordered] @{
        'Column 1' = 'Some random text'
        'Column2' = 345
        'Custom Property' = $true
    }
    Table -Hashtable $hashtable
}
$example16 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

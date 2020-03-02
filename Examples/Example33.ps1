[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example33 = Document -Name 'PScribo Example 33' {

    <#
        A keyed list combines a collection of objects by a single property (key)
        into a single table. All other properties are displayed as individual
        rows.
    #>

    $servers = @(
        [Ordered] @{ ComputerName = 'DC1'; DomainName = 'example.local'; FQDN = 'dc1.example.local'; IpAddress = '192.168.0.1' }
        [Ordered] @{ ComputerName = 'DC2'; DomainName = 'example.local'; FQDN = 'dc2.example.local'; IpAddress = '192.168.0.2' }
        [Ordered] @{ ComputerName = 'DC3'; DomainName = 'example.local'; FQDN = 'dc3.example.local'; IpAddress = '192.168.0.3' }
    )

    Table -Hashtable $servers -List -Key 'ComputerName'

    <#
        The table above, will be rendered like so:

        ComputerName DC1               DC2               DC3
        ------------ ---               ---               ---
        DomainName   example.local     example.local     example.local
        FQDN         dc1.example.local dc2.example.local dc3.example.local
        IpAddress    192.168.0.1       192.168.0.2       192.168.0.3
    #>

}
$example33 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

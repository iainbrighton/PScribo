[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example11 = Document -Name 'PScribo Example 11' {

    <#
        As the 'Table' keyword supports pipeline input, you can use the standard
        Powershell 'Sort-Object', 'Group-Object', 'Select-Object' and 'Where-Object'
        cmdlets etc. to filter, sort and/or group the input into the 'Table' cmdlet.

        The following example creates a table of services, displaying the service
        names, display names and service statuses.

        NOTE: Due the the length of the table cell content, text output may be
              truncated due to limitations/implementation of the 'Format-Table' cmdlet.
    #>
    Get-Service | Select-Object Name,DisplayName,Status | Table
}
$example11 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

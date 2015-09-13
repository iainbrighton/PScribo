Import-Module PScribo -Force;

$example11 = Document -Name 'PScribo Example 11' {
    <#
        As the 'Table' keyword supports pipeline input, you can use the standard
        Powershell Sort-Object, Group-Object, Select-Object and Where-Object cmdlets
        etc. to filter, sort and/or group the input into the 'Table' cmdlet.

        The following example creates a table of services, displaying the service
        names, display names and service statuses.
    #>
    Get-Service | Select-Object Name,DisplayName,Status | Table
}
$example11 | Export-Document -Format Html -Path ~\Desktop

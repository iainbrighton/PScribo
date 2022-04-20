[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example28 = Document -Name 'PScribo Example 28' {

    <#
        Custom styling can also be applied to individual cells within a table. To
        flag an individual object property to apply a different style, an additional
        property must be added to the object(s). Instead of adding a '__Style' property
        add a '<PropertyName>__Style' string property.

        For example, to style a service 'Name' property with a custom style, add an
        additional property called 'Name__Style' with the style to apply.

        NOTE: Individual cell styling will override any row style applied via the
              '__Style' property.

        How the properties are added does not matter. However, the Set-Style cmdlet
        can also apply a specified style to one or more object properties with
        the -Property parameter.

        The following example styles the individual table cells wherever the service
        status is not 'Running'.
    #>
    Style -Name StoppedService -Color White -BackgroundColor Firebrick

    $services = Get-Service
    $services | Where-Object { $_.Status -ne 'Running' } | Set-Style -Style 'StoppedService' -Property 'Status'

    Table -InputObject $services -Columns Name,DisplayName,Status -Headers 'Name','Display Name','State'
}
$example28 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

Import-Module PScribo -Force;

$example30 = Document -Name 'PScribo Example 30' {
    <#
        When sending PScribo output as a HTML email body, Outlook does not necessarily
        display the page styling/look-and-feel correctly. This can be suppressed by
        passing an -Options hashtable with the 'NoPageLayoutStyle' key to the
        Export-Document function.
    #>
    GlobalOption -Margin 18
    Style -Name StoppedService -Color White -BackgroundColor Firebrick

    $services = Get-Service
    $services | Where-Object { $_.Status -ne 'Running' } | Set-Style -Style 'StoppedService'
    Table -InputObject $services -Columns Name,DisplayName,Status -Headers 'Name','Display Name','State'
}
$example30 | Export-Document -Format Html -Path ~\Desktop -Options @{ NoPageLayoutStyle = $true }

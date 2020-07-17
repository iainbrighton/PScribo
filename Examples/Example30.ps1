[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example30 = Document -Name 'PScribo Example 30' {

    <#
        When sending PScribo output as a HTML email body, Outlook does not necessarily
        display the page styling/look-and-feel correctly. This can be suppressed by
        passing an -Options hashtable with the 'NoPageLayoutStyle' key to the
        Export-Document function.

        Setting this option, removes the page layout styling, rendering the contents the
        full width of the web browser (or Outlook message) window.

        NOTE: Regardless, Outlook might not render the HTML5 content correctly due its
              usage of the Internet Explorer rendering engine.
    #>
    DocumentOption -Margin 18
    Style -Name StoppedService -Color White -BackgroundColor Firebrick

    $services = Get-Service
    $services | Where-Object { $_.Status -ne 'Running' } | Set-Style -Style 'StoppedService'
    Table -InputObject $services -Columns Name,DisplayName,Status -Headers 'Name','Display Name','State'
}
$example30 | Export-Document -Format $Format -Path $Path -Options @{ NoPageLayoutStyle = $true }  -PassThru:$PassThru

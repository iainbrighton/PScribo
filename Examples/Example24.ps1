[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example24 = Document -Name 'PScribo Example 24' {

    <#
        The following example combines the creation of multiple custom styles with the
        definition of a custom table style. The custom "BlueZebra" table style is then
        applied to a table.
    #>
    Style -Name 'BlueZebraHeading' -Bold -Color 039 -Font 'Segoe UI' -Size 12
    Style -Name 'BlueZebraRow' -Color 669 -Font 'Lucida Sans Unicode' -BackgroundColor E8EDFF
    Style -Name 'BlueZebraAltRow' -Color 669 -Font 'Lucida Sans Unicode'
    TableStyle -Name 'BlueZebra' -HeaderStyle BlueZebraHeading -RowStyle BlueZebraRow -AlternateRowStyle BlueZebraAltRow -PaddingTop 4 -PaddingRight 4 -PaddingBottom 4 -PaddingLeft 4

    <#
        Create a standard table using the new "BlueZebra" table style.
    #>
    Get-Service | Select-Object -Last 10 | Table -Columns Name,DisplayName,Status -Headers Name,'Display Name','State' -Style BlueZebra
}
$example24 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

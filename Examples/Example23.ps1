[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example23 = Document -Name 'PScribo Example 23' {

    <#
        By default, PScribo adds padding to each table cell within a table. Different amounts
        of padding can be applied to the top, right, bottom and left of each table cell. The
        padding is specified in points (pt).

        NOTE: Table padding applies to both standard tables and -List view tables.

        The default padding attempts to ensure that table padding renders evenly in both
        Html and Word without adding too much space. The default values are as follows:

            Padding-Top   : 1.0pt
            Padding-Right : 4.0pt
            Padding-Bottom: 4.0pt
            Padding-Left  : 0.0pt

        The following example creates a custom table style that creates even padding around
        all table cells.

        NOTE: a -BorderWidth has been specified to demonstrate the padding but is not required!
    #>
    TableStyle -Name 'LargeGrid' -HeaderStyle Normal -RowStyle Normal -BorderWidth 1 -PaddingTop 4 -PaddingRight 4 -PaddingBottom 4 -PaddingLeft 4
    Get-Service | Select-Object -Property Name,DisplayName,Status -First 3 | Table -Style LargeGrid
}
$example23 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

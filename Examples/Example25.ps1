Import-Module PScribo -Force;

$example25 = Document -Name 'PScribo Example 25' {
    <#
        The following example combines the creation of multiple custom styles with the
        definition of a custom table style. The custom "BlueZebra" table style is then
        applied to a -List view table of the 11th to last service.

        NOTE: List view tables do not apply the -AlternateRowStyle styling. Therefore,
        the -AlternateRowStyle property only needs to be defined if the style is
        applied to both "standard" and -List view tables.
    #>
    Style -Name 'BlueZebraHeading' -Bold -Color 039 -Font 'Segoe UI' -BackgroundColor E8EDFF
    Style -Name 'BlueZebraRow' -Color 669 -Font 'Lucida Sans Unicode'
    TableStyle -Name 'BlueZebra' -HeaderStyle BlueZebraHeading -RowStyle BlueZebraRow -PaddingTop 4 -PaddingRight 4 -PaddingBottom 4 -PaddingLeft 4 -BorderWidth 1 -BorderColor E8EDFF

    <#
        Create a standard table using the new "BlueZebra" table style.
    #>
    Get-Service | Select-Object -Last 1 -Skip 10 | Table -Columns 'Name','DisplayName','Status' -Headers 'Name','Display Name','State' -ColumnWidths 25,75 -Style BlueZebra -List
}
$example25 | Export-Document -Format Html -Path ~\Desktop

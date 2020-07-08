[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example36 = Document -Name 'PScribo Example 36' {

    <#
        Headers and footers are displayed inside the page margin. To increase
        the available space, the top and bottom margins can be adjusted to
        compensate.
    #>
    DocumentOption -MarginTopAndBottom 36 -MarginLeftAndRight 54

    <#
        There are no default styles defined for headers and footers, but they
        can be styled in the standard way we style paragraphs and tables.
    #>
    Style -Name 'Header' -Size 12 -Color 0072af -Align Center -Bold
    Header -Default -IncludeOnFirstPage {
        Paragraph -Style 'Header' 'PScribo Example 36'
    }

    <#
        It is also possible to define a custom style and assign it to a
        header and/or footer.
    #>
    Style -Name 'CustomFooter' -Size 11 -Color 0072af -Align Center -Italic
    Footer -Default {
        Paragraph -Style 'CustomFooter' 'Page <!# PageNumber #!> of <!# TotalPages #!>'
    }

    Get-Service |
        Select-Object -First 25 -Property 'Name','DisplayName','Status' |
            Table -ColumnWidths 42,42,16

    PageBreak

    Get-Service |
        Select-Object -First 25 -Skip 25 -Property 'Name','DisplayName','Status' |
            Table -ColumnWidths 42,42,16

    PageBreak

    Get-Service |
        Select-Object -First 25 -Skip 50 -Property 'Name','DisplayName','Status' |
            Table -ColumnWidths 42,42,16

}
$example36 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

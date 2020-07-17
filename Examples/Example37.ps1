[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example37 = Document -Name 'PScribo Example 37' {

    DocumentOption -MarginTopAndBottom 36 -MarginLeftAndRight 54

    <#
        Create a borderless table style based on the built-in 'Normal' style.
        The 'RightAligned' style can be used to right-align the header and
        footer display.
    #>
    TableStyle -Name Borderless -HeaderStyle Normal -RowStyle Normal
    Style -Name RightAligned -Align Right

    <#
        You can create a single row split table by defining a list table with
        a single key/value pair. The property name will be displayed left-
        aligned and the property value will be right-aligned (using the style
        defined earlier).

        By default headers and footers include some additional space. This
        additional space can be removed by specifying the -NoSpace parameter.
        This is normally not needed, but an additional space is also added
        when outputting tables.
    #>
    Header -Default -IncludeOnFirstPage -NoSpace {
        $header = [Ordered] @{
            'PScribo'        = 'Example 37'
            'PScribo__Style' = 'RightAligned'
        }
        Table -HashTable $header -Style Borderless -List
    }

    <#
        You can also include the PScribo tokens in the table output.
    #>
    Footer -Default -IncludeOnFirstPage -NoSpace {
        $footer = [Ordered] @{
            "Example37.ps1"        = 'Page <!# PageNumber #!> of <!# TotalPages #!>'
            "Example37.ps1__Style" = 'RightAligned'
        }
        Table -HashTable $footer -Style Borderless -List
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
$example37 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

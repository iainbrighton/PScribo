[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example35 = Document -Name 'PScribo Example 35' {

    DocumentOption -MarginTopAndBottom 36 -MarginLeftAndRight 54

    <#
        Headers and footers can be added to a document. There are two headers
        (default and first page) and two footers (default and first page)
        that can be set. The first page header/footer is only displayed on
        the very first page, but only when defined. Default headers/footers
        are displayed on ALL other pages.

        The only components that can be used within a Header or Footer
        element are 'Paragraph' and 'Table'. Any other components will be
        ignored and display a warning.
    #>

    Header -FirstPage {

        <#
            The first page header will only be shown on the first page. If no
            first page header is defined, nothing will be displayed.
        #>
        Paragraph 'PScribo'
    }

    Header -Default {

        <#
            The 'Default' header will be displayed on all subsequent pages. If
            no default header is defined, nothing will be displayed.
        #>
        Paragraph 'PScribo Example 35'
    }

    Footer -Default -IncludeOnFirstPage {
        <#
            Headers and footers that should be displayed on all pages can be
            defined in a single pass.

            There are two tokens that can be used in the header/footer to
            insert page number references:

                <!# PageNumber #!> will insert the current page number
                <!# TotalPages #!> will inset the total number of pages

            NOTE: All formats other than Word, will only insert an approximate
                  page number. Word will reflow and recalculate the correct
                  page numbers.
        #>
        Paragraph 'Page <!# PageNumber #!> of <!# TotalPages #!>'
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
$example35 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

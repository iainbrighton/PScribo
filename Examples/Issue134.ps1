[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module D:\Repos\PScribo -Force -Verbose:$false

$issue134 = Document -Name 'PScribo Example 37' {

    ### Style copied from AsBuiltReport

    # Configure document options
        DocumentOption -EnableSectionNumbering -PageSize A4 -DefaultFont 'Segoe Ui' -MarginLeftAndRight 71 -MarginTopAndBottom 71 -Orientation Portrait

        # Configure Title & Heading Styles
        Style -Name 'Title' -Size 24 -Color '072E58' -Align Center
        Style -Name 'Title 2' -Size 18 -Color '204369' -Align Center
        Style -Name 'Title 3' -Size 12 -Color '395879' -Align Left
        Style -Name 'Heading 1' -Size 16 -Color '072E58'
        Style -Name 'Heading 2' -Size 14 -Color '204369'
        Style -Name 'Heading 3' -Size 13 -Color '395879'
        Style -Name 'Heading 4' -Size 12 -Color '958026'
        Style -Name 'NO TOC Heading 4' -Size 12 -Color '958026'
        Style -Name 'Heading 5' -Size 11 -Color '009684'
        Style -Name 'NO TOC Heading 5' -Size 11 -Color '009684'
        Style -Name 'Heading 6' -Size 10 -Color '009683'
        Style -Name 'NO TOC Heading 6' -Size 10 -Color '009683'
        Style -Name 'NO TOC Heading 7' -Size 10 -Color '00EBCD' -Italic
        Style -Name 'Normal' -Size 10 -Color '565656' -Default
        # Header & Footer Styles
        Style -Name 'Header' -Size 10 -Color '565656' -Align Center
        Style -Name 'Footer' -Size 10 -Color '565656' -Align Center
        # Table of Contents Style
        Style -Name 'TOC' -Size 16 -Color '072E58'
        # Table Heading & Row Styles
        Style -Name 'TableDefaultHeading' -Size 10 -Color 'FAFAFA' -BackgroundColor '072E58'
        Style -Name 'TableDefaultRow' -Size 10 -Color '565656'
        # Table Row/Cell Highlight Styles
        Style -Name 'Critical' -Size 10 -Color '565656' -BackgroundColor 'FEDDD7'
        Style -Name 'Warning' -Size 10 -Color '565656' -BackgroundColor 'FFF4C7'
        Style -Name 'Info' -Size 10 -Color '565656' -BackgroundColor 'E3F5FC'
        Style -Name 'OK' -Size 10 -Color '565656' -BackgroundColor 'DFF0D0'
        # Table Caption Style
        Style -Name 'Caption' -Size 10 -Color '072E58' -Italic -Align Left
        Style -Name RightAligned -Size 10 -Color '565656' -Align Right

        # Configure Table Styles
        $TableDefaultProperties = @{
            Id = 'TableDefault'
            HeaderStyle = 'TableDefaultHeading'
            RowStyle = 'TableDefaultRow'
            BorderColor = '072E58'
            Align = 'Left'
            CaptionStyle = 'Caption'
            CaptionLocation = 'Below'
            BorderWidth = 0.25
            PaddingTop = 1
            PaddingBottom = 1.5
            PaddingLeft = 2
            PaddingRight = 2
        }

        TableStyle @TableDefaultProperties -Default
        TableStyle -Id 'Borderless' -HeaderStyle Normal -RowStyle Normal -BorderWidth 0
    # DocumentOption -MarginTopAndBottom 36 -MarginLeftAndRight 54

    <#
        Create a borderless table style based on the built-in 'Normal' style.
        The 'RightAligned' style can be used to right-align the header and
        footer display.

    TableStyle -Name Borderless -HeaderStyle Normal -RowStyle Normal
    Style -Name RightAligned -Align Right
    #>

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

    PageBreak

    List {
        Item 'Apples'
        Item 'Pears'
        Item 'Bananas'
    }

    NumberStyle -Id 'LeftRoman' -Format Roman -Align Left
    NumberStyle -Id 'RightRoman' -Format Roman -Align Right
    List -Numbered -NumberStyle RightRoman {
        Item 'Apples'
        List -Numbered -NumberStyle RightRoman {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List -Numbered -NumberStyle RightRoman {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }
}
$issue134 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

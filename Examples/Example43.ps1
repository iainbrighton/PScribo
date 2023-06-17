[CmdletBinding()]
param (
    [System.String[]] $Format = 'Word',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example43 = Document -Name 'PScribo Example 43' {

    <#
        PScribo provides 3 built-in number styles that replicate the standard Html options and
        the built-in Microsoft Word defaults; 'Number', 'Letter' and 'Roman'. The default number
        styles display the number in lowercase, right-aligned and terminated with a period '.'.

        It is possible to define your own number styles or override the built-in styles. This
        provides options to change the casing and/or alignment of the list numbers.

        NOTE: Html numbered lists only support the default '.' number style terminator/suffix. The
              use of custom number style terminators/suffixes is not supported.

        NOTE: Html numbered/unordered lists do not support alignment.

        For example, to override the built-in number styles ensuring that they are rendered in
        uppercase and terminated with a parenthesis:
    #>
    NumberStyle -Id 'Number' -Format Number -Uppercase -Suffix ')'
    NumberStyle -Id 'Letter' -Format Letter -Uppercase -Suffix ')'
    NumberStyle -Id 'Roman' -Format Roman -Uppercase -Suffix ')'

    <#
        To align the number to the left margin, override or define your own number style with the
        '-Align' parameter.

        NOTE: The default number style has been changed so we define a 'RightRoman' to mimic the
              built-in/default settings (without redefining 'Roman' again!).
    #>
    NumberStyle -Id 'LeftRoman' -Format Roman -Align Left
    NumberStyle -Id 'RightRoman' -Format Roman -Align Right

    <#
        Output right aligned (the default) lists for comparison
    #>
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

    <#
        Output left aligned lists for comparison

        NOTE: Html numbered lists only support the default '.' number style terminator/suffix. The
              use of custom number style terminators/suffixes is not supported.
    #>
    List -Numbered -NumberStyle LeftRoman {
        Item 'Apples'
        List -Numbered -NumberStyle LeftRoman {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List -Numbered -NumberStyle LeftRoman {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }
}
$example43 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

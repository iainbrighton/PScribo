[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example42 = Document -Name 'PScribo Example 42' {

    <#
        Multi-level numbered lists can be created by nesting one or more 'List' keywords in
        combination with the '-Numbered' parameter.

        NOTE: A 'List' defaults to a bulleted list by default, so the '-Numbered' switch needs
              to be specified at each level - where required.
    #>
    List -Numbered {
        Item 'Apples'
        List -Numbered {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List -Numbered {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }

    <#
        Like bullet lists, each 'List' can have its own number style defined.

        NOTE: Word does not support a mixture of bullet/number formats at the same level within a
              list. Therefore, only the first list type will be rendered at each level - in this
              example the 'Letter' style will be used for the second nested numbered list.
    #>
    List -Numbered {
        Item 'Apples'
        List -Numbered -NumberStyle Letter {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List -Numbered -NumberStyle Roman {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }

}
$example42 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

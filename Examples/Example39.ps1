[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example39 = Document -Name 'PScribo Example 39' {

    <#
        Bulleted and numbered lists are supported. A list is defined by the 'List'
        keyword and can contain one or more items.

        A simple single-level list can be defined with the '-Item' parameter and
        passing an array of [string[]].
    #>
    List -Item 'Apples','Oranges','Bananas'

    <#
        Lists default to a bulleted list, but a numbered list can be created wtih
        the '-Numbered' parameter.
    #>
    List -Item 'Apples','Oranges','Bananas' -Numbered

    <#
        A list can also be created using a script block and nesting one or 'Item'
        within it.
    #>
    List {
        Item 'Apples'
        Item 'Bananas'
        Item 'Oranges'
    }

    <#
        Multi-level lists can be created by nesting the 'List' within a script
        block.

        NOTE: nested 'List' objects can only be nested after the inclusion of
              an 'Item'.
    #>
    List {
        Item 'Apples'
        List {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }

    <#
        Multi-level numbered lists can also be created with the '-Numbered' parameter.
    #>
    List -Numbered {
        Item 'Apples'
        List {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }

}
$example39 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

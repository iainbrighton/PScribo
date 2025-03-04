[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example41 = Document -Name 'PScribo Example 41' {

    <#
        Multi-level bullet lists can be created by nesting one or more 'List' keywords.

        NOTE: There must be an 'Item' keyword defined before a nested 'List' can be used.
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
        Each 'List' can have its own bullet style defined.

        NOTE: Word does not support a mixture of bullet/number formats at the same level within a
              list. Therefore, only the first list type will be rendered at each level - in this
              example the 'Disc' style will be used.

        NOTE: Html output does not support the 'Dash' bullet style. Dashes will be rendered using
              the the web broswer's defaults.
    #>
    List -BulletStyle Square {
        Item 'Apples'
        List -BulletStyle Disc {
            Item 'Jazz'
            Item 'Granny smith'
            Item 'Pink lady'
        }
        Item 'Bananas'
        Item 'Oranges'
        List -BulletStyle Dash {
            Item 'Jaffa'
            Item 'Tangerine'
            Item 'Clementine'
        }
    }

}
$example41 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

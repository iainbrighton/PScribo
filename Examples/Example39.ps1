[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

# 'List' single-level bullet lists
$example39 = Document -Name 'PScribo Example 39' {

    <#
        A bulleted list is defined by the 'List' keyword and can contain one or more items.

        A simple single-level list can be defined with the '-Item' parameter and passing an
        array of strings ([string[]]).
    #>
    List -Item 'Apples','Oranges','Bananas'

    <#
        A list can also be created using a script block and nesting one or 'Item' keywords
        within it.
    #>
    List {
        Item 'Apples'
        Item 'Bananas'
        Item 'Oranges'
    }

    <#
        Bullet styles can be applied to a list, e.g. 'Dash', 'Circle', 'Disc' and 'Square'. If
        not specified, the bullet list defaults to the 'Disc' style.
    #>
    List -BulletStyle Square {
        Item 'Apples'
        Item 'Bananas'
        Item 'Oranges'
    }

    <#
        Formatting styles can be applied to all items in a list.
    #>
    List -Style Caption {
        Item 'Apples'
        Item 'Bananas'
        Item 'Oranges'
    }

    <#
        Styles can be applied to indiviual items in a list.
    #>
    List {
        Item 'Apples'
        Item 'Bananas' -Style Caption
        Item 'Oranges'
    }

    <#
        Inline styles can also be applied to indiviual items in a list.
    #>
    List {
        Item 'Apples' -Bold
        Item 'Bananas' -Italic
        Item 'Oranges' -Color Firebrick
    }

}
$example39 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

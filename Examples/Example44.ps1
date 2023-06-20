[CmdletBinding()]
param (
    [System.String[]] $Format = 'Word',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example44 = Document -Name 'PScribo Example 44' {

    <#
        Custom numbered lists are registered with the 'NumberStyle' keyword, but only the Word and
        Text plugins are supported. All other plugins will render the number as a decimal (using the
        'Number' format).

        Custom number lists can contain any wording and punctuation you require.

        NOTE: The '-Uppercase' and '-Suffix' parameters are ignored so you need to include any suffix
              in the number format definition.

        The '%' token is used to denote where the number will be placed. To include leading zeroes,
        use multiple '%' tokens, e.g. 'ab%%' for ab01, ab02 and 'XYZ-%%%' for XYZ-001, XYZ-002, etc..
    #>
    NumberStyle -Id 'CustomNumberStyle' -Custom 'xYz-%%%.' -Indent 1500 -Hanging 200 -Align Left

    <#
        Output list using the 'Custom' number style
    #>

    List -Numbered -NumberStyle CustomNumberStyle -Item 'Apples','Bananas','Oranges'

}
$example44 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

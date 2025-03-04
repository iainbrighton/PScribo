[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example40 = Document -Name 'PScribo Example 40' {

    <#
        Numbered lists are supported and can be specified with the '-Numbered' parameter.
    #>
    List -Item 'Apples','Oranges','Bananas' -Numbered

    <#
        Multiple number styles are available: 'Number', 'Letter' and 'Roman'. If not specified,
        a numbered list will default to the 'Number' style. You can specify the required number
        format with the '-NumberStyle' parameter.
    #>
    List -Item 'Apples','Oranges','Bananas' -Numbered -NumberStyle Letter

    <#
        Multiple number styles are available: 'Number', 'Letter' and 'Roman'. If not specified,
        a numbered list will default to the 'Number' style. You can specify the required number
        format with the '-NumberStyle' parameter.
    #>
    List -Numbered -NumberStyle Roman {
        Item 'Apples'
        Item 'Bananas'
        Item 'Oranges'
    }

}
$example40 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

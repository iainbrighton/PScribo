[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example5 = Document -Name 'PScribo Example 5' {

    <#
        Rather than specifying styling options on individual paragraphs, PScribo supports defining (or
        overriding the default styles) with your own. Creating your own style is easy - just use the
        'Style' cmdlet.

        The following command creates a new style called 'Funky':
    #>
    Style -Name 'Funky' -Font Arial -Size 16 -Color Pink -Align Right

    <#
        The style can then easily be applied to a paragraph by specifying the -Style parameter.
    #>
    Paragraph 'This paragraph is styled with the "Funky" style!' -Style Funky

    <#
        If no style is specified, the default 'Normal' is used. You can override the default style
        by defining your own style with the same name:
    #>
    Style -Name 'Normal' -Font Tahoma -Size 12 -Color 000
    Paragraph 'This paragraph will be styled with the custom "Normal" style defined earlier.'
}
$example5 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

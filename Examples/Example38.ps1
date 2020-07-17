[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example38 = Document -Name 'PScribo Example 38' {

    <#
        Paragraphs are comprised of one or more text "runs". A text run is a body
        of text that shares a common set of properties, e.g. font, weight and/or
        color. Using multiple text runs allows full customisation of a paragraph's
        formatting.

        As legacy paragraphs did not support inline styling, they are implemented
        as a single paragraph with a single text run. If you don't need inline
        styling, you can continue to use the legacy 'Paragraph' parameters.

        The new paragraph formatting requires a script block:
    #>
    Paragraph {

        <#
            Each individual body of text is defined using the 'Text' keyword.
            Consecutive text runs are automatically separated by a space, unless
            the '-NoSpace' switch parameter is specified.
        #>
        Text 'This is the first paragraph text run.'
        Text 'This is the second paragraph text run that will be appended to the previous text run.'
    }

    Blankline

    <#
        Styling can be applied at the paragraph level and overridden where
        necessary on each text run.
    #>
    Style -Name 'Custom' -Font 'Times New Roman' -Size 10
    Paragraph -Style 'Custom' {

        Text 'This is the first paragraph text run.'
        Text 'This is the second paragraph text run that will be appended to the previous text run.' -Style Caption
    }

    Blankline

    Paragraph -Style 'Custom' {

        <#
            Inline styling can also be applied to a text run without having to
            define a style.
        #>
        Text 'This is the first paragraph text run.'
        Text 'This is the second paragraph text run that' -Style Caption
        Text 'will' -Bold -Italic -Underline -Color 'Firebrick' -NoSpace
        Text ' be appended to the previous text run.' -Style Caption
    }

}
$example38 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example6 = Document -Name 'PScribo Example 6' {

    <#
       You can create a document structure with the 'Section' cmdlet. This is the only PScribo
       cmdlet that supports nesting of script blocks.

       Sections are used to create a hierarchy within the document that are used when generating
       a table of contents. If you have no requirement for a table of contents, then you may
       not need to use the 'Section' keyword.

       You can define sections with the 'Section' keyword, a name and script block. You can nest
       sections within sections.
    #>
    Section -Name 'First Section' -ScriptBlock {
        Paragraph 'This section will not have any explicit styling applied.'
    }

    <#
        Styles can also be applied to Sections with the -Style parameter. PScribo defines the
        following heading styles: Heading1, Heading2 and Heading3. Just like the built-in
        "Normal" style you can define your own styles or override the default ones.
    #>
    Section -Name 'Second Styled Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section heading will be styled with the built-in "Heading1" style.'
    }

    Section -Name 'Third Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This paragraph is nested within "Third Section".'

        Section 'Sub Section' -Style Heading2 {
            Paragraph 'This paragraph is nested beneath the "Third Section\Sub Section".'
        }
    }
}
$example6 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

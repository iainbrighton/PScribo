[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example9 = Document -Name 'PScribo Example 9' {
    DocumentOption -EnableSectionNumbering -MarginTopAndBottom 72 -MarginLeftAndRight 54
    TOC -Name 'Table of Contents'

    <#
       PScribo supports multiple types of "breaks" within a document; line breaks, page
       breaks and blank lines. This can be used to segregate content within a document.

       NOTE: not all breaks are rendered by PScribo plugins in the same manner. For example,
             only Word output truly supports page breaks as there is no concept of pagination
             in html, xml or text output.

       The 'PageBreak' cmdlet is used to insert a page break. The following page break
       inserts a break between the table of content and the first section.
    #>
    PageBreak

    Section -Name 'First Section' -ScriptBlock {
        Paragraph 'This section should be labeled as "1 First Section".'
    }

    <#
        A 'LineBreak' inserts a horizontal line across the page. This can be used to
        separate document contents without requiring a page break.
    #>
    LineBreak

    Section -Name 'Second "Styled" Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section should be labeled as "2 Second Styled Section".'
    }

    Section -Name 'Third "Styled" Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section should be labeled as "3 Third Section".'
        <#
            Blank lines can be inserted into a document with the 'BlankLine' cmdlet.
            Specifying a single 'BlankLine' command is equivalent to creating an
            empty paragraph.
        #>
        BlankLine

        Section 'Sub Section' -Style Heading2 {
            <#
                Mulitple blank lines can be inserted by utilising the -Count parameter. This
                can be useful when creating a cover or title page.
            #>
            BlankLine -Count 5
            Paragraph 'This section should be labeled as "3.1 Sub Section".'
        }
    }
}
$example9 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

Import-Module PScribo -Force;

$example8 = Document -Name 'PScribo Example 8' {
    GlobalOption -EnableSectionNumbering -MarginTopAndBottom 72 -MarginLeftAndRight 54

    <#
       A table of contents can be inserted into the document with the 'TOC' cmdlet. The text
       that is displayed is specified using the -Name parameter. If this is not specified, the
       default 'Contents' is displayed.

       The TOC cmdlet will create (hyper)links between the table of content and all sections
       defined within the document. NOTE: links are only supported in Html and Word output.
    #>
    TOC -Name 'Table of Contents'

    Section -Name 'First Section' -ScriptBlock {
        Paragraph 'This section should be labeled as "1 First Section".'
    }

    <#
        You can exclude sections from the TOC by specifying the -ExcludeFromTOC parameter
        on the 'Section' cmdlet.

        NOTE: Microsoft Word creates the table of contents based on the style name. Therefore,
        to include a section in a Word TOC it must be styled with a "Heading*" style name. In
        addition, if the -ExcludeFromTOC is specified with a style name of "Heading*" then
        Microsoft Word will include it - regardless.
    #>
    Section -Name 'Second "Styled" Section' -Style Heading1 -ExcludeFromTOC -ScriptBlock {
        Paragraph 'This section will be excluded from the table of contents.'
    }

    Section -Name 'Third "Styled" Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section should be labeled as "3 Third Section".'

        Section 'Sub Section' -Style Heading2 {
            Paragraph 'This section should be labeled as "3.1 Sub Section".'
        }
    }
}
$example8 | Export-Document -Format Html -Path ~\Desktop

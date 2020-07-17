[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example8 = Document -Name 'PScribo Example 8' {
    DocumentOption -EnableSectionNumbering -MarginTopAndBottom 72 -MarginLeftAndRight 54

    <#
       A table of contents can be inserted into the document with the 'TOC' cmdlet. The text
       that is displayed is specified using the -Name parameter. If this is not specified, the
       default 'Contents' is displayed.

       The TOC cmdlet will create (hyper)links between the table of content and all sections
       defined within the document. NOTE: links are only supported in Html and Word output.
    #>
    TOC -Name 'Table of Contents'

    <#
        Microsoft Word creates the table of contents based on the name of the style applied. To
        include a section in a Word TOC it must be styled with a "Heading*" style name. In this
        instance, Microsoft Word will NOT include it as there no "Heading*" style applied.
    #>
    Section -Name 'First Section' -ScriptBlock {
        Paragraph 'This section should be labeled as "1 First Section".'
    }

    <#
        You can exclude sections from the TOC by specifying the -ExcludeFromTOC parameter
        on the 'Section' cmdlet.

        NOTE: As this section is styled with the "Heading1" style, Microsoft Word will still
              include it in the TOC, regardless of the -ExcludeFromTOC switch parameter.
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
$example8 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

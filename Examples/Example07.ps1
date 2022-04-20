[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example7 = Document -Name 'PScribo Example 7' {

    <#
       Sections support automatic numbering, i.e. PScribo will automatically generate the section
       numbers/levels based on the nesting. To turn this on, use the 'DocumentOption' cmdlet.
    #>
    DocumentOption -EnableSectionNumbering -MarginTopAndBottom 72 -MarginLeftAndRight 54

    Section -Name 'First Section' -ScriptBlock {
        Paragraph 'This section should be labeled as "1 First Section".'
    }

    Section -Name 'Second "Styled" Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section should be labeled as "2 Second Styled Section".'
    }

    Section -Name 'Third "Styled" Section' -Style Heading1 -ScriptBlock {
        Paragraph 'This section should be labeled as "3 Third Section".'

        Section 'Sub Section' -Style Heading2 {
            Paragraph 'This section should be labeled as "3.1 Sub Section".'
        }
    }
}
$example7 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

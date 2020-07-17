[CmdletBinding()]
param (
    [System.String[]] $Format = 'Word',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example32 = Document -Name 'PScribo Example 32' {

    <#
        A document has a default 'Portrait' page orientation. This can always be overridden using the
        `DocumentOption -Orientation Landscape` function if required. However, this changes the page orientation
        for all sebequent pages.
    #>
    DocumentOption -Orientation Portrait

    Section -Style 'Heading1' -Name 'Default Page Orientation' {

        <#
            You can specify an orientation on each document root `Section`. When no orientation is specified it will be
            displayed in the document's "current" orientation. If no orientation has been explicitly set, then the
            document's default page orientation is used.
        #>
        Paragraph 'This section will be displayed in "Portrait" as no orientation has been set'
    }

    Section -Style 'Heading1' -Name 'Landscape Section' -Orientation Landscape {

        Paragraph 'This paragraph will be displayed in "Landscape" as the orientation has been specified.'
    }

    Section -Style 'Heading1' -Name 'Continuous Orientation' {

        <#
            When the orientation is changed, it is changed for all following sections - unless the orientation is
            explicitly reverted.
        #>
        Paragraph 'This paragraph will be displayed in "Landscape" as the orientation has not been specified, but has been changed by a previous section.'

        Section -Style 'Heading2' -Name 'Orientation Warning' -Orientation Portrait {

            Paragraph 'This paragraph will be displayed in "Landscape" as orientation can only be set at document-level section blocks (a warning will also be shown!).'
        }
    }

    Section -Style 'Heading1' -Name 'Portrait Section' -Orientation Portrait {

        Paragraph 'This paragraph will be displayed in "Portrait" again as the orientation has been specified.'
    }

}
$example32 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

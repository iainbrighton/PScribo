[CmdletBinding()]
param (
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example2 = Document -Name 'PScribo Example 2' {
    Paragraph 'This is an example that adds a single paragraph to the document that will be styled with the default style.'
}

<#
    Once we have a reference to a PScribo document, it can be exported with the Export-Document cmdlet.
    The Export-Document cmdlet can export to one or more formats: Word, Html, Text and Html are included
    by default.

    This command converts/exports the $example2 document to a text formatted file named 'PScribo Example 2.txt' on the desktop.
#>
Export-Document -Document $example2 -Format Text -Path ~\Desktop

<#
    This command converts/exports the $example2 document to a html formatted file named 'PScribo Example 2.html' on the desktop.
#>
Export-Document -Document $example2 -Format Html -Path ~\Desktop

<#
    This command exports the $example2 document to both Word formatted 'PScribo Example 2.docx' and html formatted
    'PScribo Example 2.html' files on the desktop, via the pipeline.

    NOTE: This will overwrite the 'PScribo Example 2.html' created above without warning.
#>
$example2 | Export-Document -Format Word,Html -Path ~\Desktop -PassThru:$PassThru

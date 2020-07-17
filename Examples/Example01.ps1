[CmdletBinding()]
param (
    [System.String] $Path = '~\Desktop\Example1.CliXml',
    [System.Management.Automation.SwitchParameter] $PassThru
)

<#
    From Powershell v3 onwards, the module should not need to be explicitly imported. It is included
    here to avoid any ambiguity.
#>
Import-Module PScribo -Force -Verbose:$false

<#
    We create a PScribo document with the 'Document' cmdlet. Inside the supplied script block, you can include any
    standard Powershell code as well as additional PScribo content, such as paragraphs. The following code creates a
    new blank document named 'PScribo Example 1', storing it in the $example1 variable.

    When documents are exported, they are exported using the name of the document appended with the extension of the
    plugin. For example, exporting 'PScribo Example 1' to a Word document results in a 'PScribo Example 1.docx' file.
#>

$example1 = Document -Name 'PScribo Example 1' {
    <#
        This creates a single paragraph within the 'Example 1' document
    #>
    Paragraph 'This is an example that adds a single paragraph to the document that will be styled with the default style.'
}

<#
    If needed the PScribo document stored in variable $example1 could be exported via the Export-CliXml cmdlet. This would
    permit the creation of a document on one machine, and converted into an external document format on another. The following
    command will export the 'PScribo Example 1' document to a file named Example1.CliXml on the desktop.
#>
$example1 | Export-Clixml -Depth 128 -Path $Path -Force

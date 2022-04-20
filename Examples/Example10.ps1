[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example10 = Document -Name 'PScribo Example 10' {

    <#
        PScribo supports insertion of tables into a document. In it's simplest form, a
        collection of objects can be added to a document by using the 'Table' cmdlet. By
        default, a table is created the full width (100%) of the document, with each
        column width being equally distributed.

        The following examples creates a table of all properties on all services.

        NOTE: Due the the vast number of properties, the table will not render in any
              meaningful way! Text output may also be truncated due to limitations/
              implementation of the 'Format-Table' cmdlet.
    #>
    Table -InputObject (Get-Service | Select-Object -First 100)
}
$example10 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

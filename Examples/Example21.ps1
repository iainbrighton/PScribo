[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example21 = Document -Name 'PScribo Example 21' {

    <#
        Styling can also be applied to tables - just like Paragraphs and Sections. Tables
        support additional styling options such as borders and cell padding options.

        NOTE: There is a built-in "TableDefault" style. You can override the default table
              style just like overriding the default "Normal" paragraph style.

        Each table style also requires three styles; one applied to the header row, one
        applied to each table row and an optional alternating row style.

        Before table styles can be defined, the individual styles must already be defined
        in the document via the 'Style' cmdlet/keyword.

        NOTE: The "TableDefault" style uses a style called "TableDefaultHeading" for the
              header row, a style called "TableDefaultRow" for the row style and a style
              called "TableDefaultAltRow" for the alternating row style.

        The following defines a very simple table style named "Basic" that uses the "Normal"
        style for the header and all other rows, e.g. no styling!

        NOTE: When -AlternateRowStyle is not specified, the -RowStyle property value is
              used for -AlternateRowStyle property.
    #>
    TableStyle -Name 'Basic' -HeaderStyle Normal -RowStyle Normal
    Get-Service | Select-Object -Property Name,DisplayName,Status -First 3 | Table -Style Basic
}
$example21 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

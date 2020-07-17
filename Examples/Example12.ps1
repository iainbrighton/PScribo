[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example12 = Document -Name 'PScribo Example 12' {
    <#
        The 'Table' cmdlet supports filtering of object properties by using the -Columns
        parameter. The -Columns parameter specifies the object properties you wish to
        display in the table, in the order required.

        NOTE: The -Columns parameter is required if you wish to override the values
              displayed in column headings.

        The following example displays the Name, DisplayName and Status properties of
        all services sent down the pipeline, ignoring all other properties.

        NOTE: The columns are displayed in the exact order that they are listed in
              the -Columns parameter.

        NOTE: Due the the length of the table cell content, text output may be
              truncated due to limitations/implementation of the 'Format-Table' cmdlet.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status
}
$example12 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example13 = Document -Name 'PScribo Example 13' {

    <#
        To override the column headers, the -Columns and -Headers parameters can be used
        together.

        The -Headers parameter specifies the heading names to apply to the corresponding
        (positional) -Columns parameter. This permits you to include spaces in the headers
        and/or change the properties' display values.

        The following example overrides the column names with the values supplied in the
        -Headers parameter - a space is introduced in the 'DisplayName' property and the
        'Status' property is displayed as 'State' instead.

        NOTE: Due the the length of the table cell content, text output may be
              truncated due to limitations/implementation of the 'Format-Table' cmdlet.
    #>
    Get-Service | Table -Columns Name,DisplayName,Status -Headers Name,'Display Name',State
}
$example13 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru -Options @{ TextWidth = 160 }

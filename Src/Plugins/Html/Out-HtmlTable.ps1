function Out-HtmlTable
{
<#
    .SYNOPSIS
        Output formatted Html <table> from PScribo.Table object.

    .NOTES
        One table is output per table row with the -List parameter.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        $formattedTable = Get-HtmlTable -Table $Table
        [ref] $null = $tableBuilder.Append($formattedTable)
        return $tableBuilder.ToString()
    }
}

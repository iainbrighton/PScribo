function ConvertTo-PScriboPreformattedTable
{
<#
    .SYNOPSIS
        Creates a formatted table based upon table type for plugin output/rendering.

    .NOTES
        Maintains backwards compatibility with other plugins that do not require styling/formatting.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        if ($Table.IsKeyedList)
        {
            Write-Output -InputObject (ConvertTo-PScriboFormattedKeyedListTable -Table $Table)
        }
        elseif ($Table.IsList)
        {
            Write-Output -InputObject (ConvertTo-PScriboFormattedListTable -Table $Table)
        }
        else
        {
            Write-Output -InputObject (ConvertTo-PScriboFormattedTable -Table $Table)
        }
    }
}

function ConvertTo-PSObjectKeyedListTable
{
<#
    .SYNOPSIS
        Converts a PScribo.Table to a [PSCustomObject] collection representing a keyed list table.
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
        $listKey = $Table.ListKey
        $rowHeaders = $Table.Rows | Select-Object -ExpandProperty $listKey
        $columnHeaders = $Table.Rows |
                            Select-Object -First 1 -Property * -ExcludeProperty '*__Style' |
                                ForEach-Object { $_.PSObject.Properties.Name } |
                                    Where-Object { $_ -ne $listKey }

        foreach ($columnHeader in $columnHeaders)
        {
            $psCustomObjectParams = [Ordered] @{
                $listKey = $columnHeader
            }
            foreach ($rowHeader in $rowHeaders)
            {
                $psCustomObjectParams[$rowHeader] = $Table.Rows |
                    Where-Object { $_.$listKey -eq $rowHeader } |
                        Select-Object -ExpandProperty $columnHeader
            }
            $psCustomObject = [PSCustomObject] $psCustomObjectParams
            Write-Output -InputObject $psCustomObject
        }
    }
}

function Get-TableEmptyColumn {
    <#
    .SYNOPSIS
        Identifies columns in a table that contain only empty or whitespace values.

    .DESCRIPTION
        The Get-TableEmptyColumn function analyzes a table (array of PSCustomObjects) and identifies
        any columns that contain only empty or whitespace values. This can be useful for table cleanup
        or analysis purposes.

    .PARAMETER InputObject
        An array of PSCustomObjects representing the table data to analyze.

    .EXAMPLE
        $data | Get-TableEmptyColumn
        Returns an array of column names that contain only empty or whitespace values.

    .EXAMPLE
        Get-TableEmptyColumn -InputObject $data
        Same as above but using parameter syntax instead of pipeline.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNull()]
        [object]$InputObject
    )

    begin {
        Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f 'Input')
        $accumulated = [System.Collections.ArrayList]@()
    }

    process {
        if ($InputObject -is [string]) {
            throw "Input must be an array of PSObjects"
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            foreach ($item in $InputObject) {
                if ($null -ne $item) {
                    [void]$accumulated.Add($item)
                }
            }
        }
        else {
            [void]$accumulated.Add($InputObject)
        }
    }

    end {
        # Convert accumulated to array for Find-EmptyTableColumn
        $array = @($accumulated)
        $emptyColumns = Find-EmptyTableColumn -InputObject $array
        if ($null -ne $emptyColumns -and @($emptyColumns).Length -gt 0) {
            Write-PScriboMessage -Message ($localized.EmptyColumnsFound -f @($emptyColumns).Length, 'Input', ($emptyColumns -join ', '))
        } else {
            Write-PScriboMessage -Message ($localized.NoEmptyColumnsFound -f 'Input')
        }
        return $emptyColumns
    }
}

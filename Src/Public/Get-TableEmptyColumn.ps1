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
        [AllowEmptyCollection()]
        [object[]]$InputObject
    )

    begin {
        Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f 'Input')
        $accumulated = [System.Collections.ArrayList]@()
    }
    
    process {
        [void]$accumulated.Add($InputObject)
    }
    
    end {
        $emptyColumns = Find-EmptyTableColumn -InputObject $accumulated
        if ($emptyColumns.Count -gt 0) {
            Write-PScriboMessage -Message ($localized.EmptyColumnsFound -f $emptyColumns.Count, 'Input', ($emptyColumns -join ', '))
        } else {
            Write-PScriboMessage -Message ($localized.NoEmptyColumnsFound -f 'Input')
        }
        return $emptyColumns
    }
} 
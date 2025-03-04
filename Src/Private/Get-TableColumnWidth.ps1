function Get-TableColumnWidth {
    [CmdletBinding()]
    [OutputType([uint16[]])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject]
        $InputObject,

        [Parameter()]
        [ValidateScript({
            foreach ($value in $_.Values) {
                if ($value -isnot [int] -or $value -lt 1 -or $value -gt 100) {
                    throw "Column widths must be integers between 1 and 100."
                }
            }
            return $true
        })]
        [System.Collections.Hashtable]
        $CustomWidths
    )
    process {
        if ($CustomWidths) {
            $totalWidth = ($CustomWidths.Values | Measure-Object -Sum).Sum
            if ($totalWidth -gt 100) {
                throw "*exceeds the total available width*"
            }
        }

        # Check if InputObject is an array or collection, and use only the first item
        if ($null -eq $InputObject) {
            throw "InputObject cannot be null."
        }

        # Validate input type
        if ($InputObject -is [string] -or $InputObject -is [ValueType]) {
            throw "Input must be an object with properties."
        }

        # Get first object if array
        if ($InputObject -is [System.Collections.IEnumerable] -and
            $InputObject -isnot [string]) {
            $InputObject = $InputObject[0]
        }

        # Get properties from the object
        $properties = @($InputObject.PSObject.Properties)
        if ($properties.Count -eq 0) {
            throw "Input object has no properties."
        }

        $columnCount = $properties.Count
        $remainingWidth = 100

        Write-PScriboMessage -Message "Initial state: Total width = $remainingWidth, Column count = $columnCount"

        # Validate custom widths reference existing properties
        if ($CustomWidths) {
            foreach ($key in $CustomWidths.Keys) {
                if (-not ($properties | Where-Object { $_.Name -ieq $key })) {
                    throw "Custom width specified for non-existent column '$key'."
                }
            }
        }

        # Special case: single column
        if ($columnCount -eq 1) {
            Write-PScriboMessage -Message "Single column detected, using full width."
            if ($CustomWidths) {
                $customWidthKey = $CustomWidths.Keys | Where-Object { $_ -ieq $properties[0].Name } | Select-Object -First 1
                if ($customWidthKey) {
                    $result = [uint16[]]@([uint16]$CustomWidths[$customWidthKey])
                } else {
                    $result = [uint16[]]@([uint16]100)
                }
            } else {
                $result = [uint16[]]@([uint16]100)
            }
            return ,$result
        }

        # Initialize widths as ordered dictionary
        $widths = [ordered]@{}

        # Get all column names
        $columnNames = @($properties.Name)

        # Apply custom widths first
        $remainingWidth = 100
        if ($CustomWidths) {
            foreach ($column in $CustomWidths.Keys) {
                if ($columnNames -contains $column) {
                    $widths[$column] = $CustomWidths[$column]
                    $remainingWidth -= $CustomWidths[$column]
                }
            }
        }

        # If all columns have custom widths but total is less than 100%, distribute remaining width evenly
        if ($CustomWidths -and $CustomWidths.Count -eq $columnCount -and $remainingWidth -gt 0) {
            $extraPerColumn = [math]::Floor($remainingWidth / $columnCount)
            $leftover = $remainingWidth - ($extraPerColumn * $columnCount)

            foreach ($column in $columnNames) {
                $widths[$column] += $extraPerColumn
            }

            # Add any remaining width to the last column
            if ($leftover -gt 0) {
                $lastColumn = $columnNames[-1]
                $widths[$lastColumn] += $leftover
            }
        }
        # Otherwise calculate even distribution for remaining columns
        else {
            $remainingColumns = @($columnNames | Where-Object { $_ -notin $widths.Keys })
            if ($remainingColumns.Count -gt 0) {
                $evenWidth = [math]::Floor($remainingWidth / $remainingColumns.Count)
                $leftover = $remainingWidth - ($evenWidth * $remainingColumns.Count)

                foreach ($column in $remainingColumns) {
                    $widths[$column] = $evenWidth
                }

                # Add any remaining width to the last column
                if ($leftover -gt 0) {
                    $lastColumn = $remainingColumns[-1]
                    $widths[$lastColumn] += $leftover
                }
            }
        }

        # Return array of widths in column order
        return @($columnNames | ForEach-Object { $widths[$_] })
    }
}

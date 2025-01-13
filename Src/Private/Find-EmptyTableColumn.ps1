# Identifies columns in a table that contain only empty or whitespace values
function Find-EmptyTableColumn {
    [CmdletBinding()]
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [AllowEmptyCollection()]
        [object[]]$InputObject
    )

    # Return empty array for empty input
    if ($InputObject.Count -eq 0) {
        Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f 'Empty Input')
        return [string[]]@()
    }

    # Validate input is not null
    if ($null -eq $InputObject) {
        throw "Input array cannot be null."
    }

    # Validate input type
    if ($InputObject -is [string] -or $InputObject[0] -is [string]) {
        throw "Input must be an object or array of objects."
    }

    Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f 'Input')

    # Get all unique properties from all objects
    $properties = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($obj in $InputObject) {
        # Skip null objects
        if ($null -eq $obj) { continue }

        # Convert non-PSObject inputs to PSObject if possible
        try {
            if ($obj -isnot [PSObject] -and $obj -isnot [PSCustomObject]) {
                $obj = [PSObject]::new($obj)
            }

            # Get properties from this object
            $objProps = @($obj.PSObject.Properties | Select-Object -ExpandProperty Name)
            foreach ($prop in $objProps) {
                $properties.Add($prop) | Out-Null
            }
        }
        catch {
            # Skip objects that can't be converted or accessed
            continue
        }
    }

    # If no properties found, return empty array
    if ($properties.Count -eq 0) {
        Write-PScriboMessage -Message ($localized.NoEmptyColumnsFound -f 'Input')
        return [string[]]@()
    }

    Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f "Found $($properties.Count) properties")

    $emptyColumns = [System.Collections.Generic.List[string]]::new()

    foreach ($prop in $properties) {
        $isBlank = $true
        foreach ($obj in $InputObject) {
            # Skip null objects
            if ($null -eq $obj) { continue }

            try {
                # Convert non-PSObject inputs to PSObject if possible
                if ($obj -isnot [PSObject] -and $obj -isnot [PSCustomObject]) {
                    $obj = [PSObject]::new($obj)
                }

                # Skip objects that don't have this property
                if (-not ($obj.PSObject.Properties.Name -contains $prop)) { continue }

                $value = $obj.$prop
                if (![string]::IsNullOrWhiteSpace($value)) {
                    $isBlank = $false
                    break
                }
            }
            catch {
                # Skip objects that can't be converted or accessed
                continue
            }
        }
        if ($isBlank) {
            Write-PScriboMessage -Message ($localized.EmptyColumnsFound -f 1, 'Input', $prop)
            $emptyColumns.Add($prop)
        }
    }

    if ($emptyColumns.Count -gt 0) {
        Write-PScriboMessage -Message ($localized.EmptyColumnsFound -f $emptyColumns.Count, 'Input', ($emptyColumns -join ', '))
    } else {
        Write-PScriboMessage -Message ($localized.NoEmptyColumnsFound -f 'Input')
    }

    # Always return as array type, even for single items
    return [string[]]($emptyColumns.ToArray())
}

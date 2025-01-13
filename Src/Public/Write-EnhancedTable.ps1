function Write-EnhancedTable {
    <#
    .SYNOPSIS
        Creates a table with enhanced features like automatic column width calculation and empty column removal.

    .DESCRIPTION
        The Write-EnhancedTable function extends the core Table function with additional features:
        - Automatic removal of empty columns
        - Custom column width calculations with intelligent distribution
        - Metadata about removed columns

        This function wraps the core Table function while maintaining all its original functionality.

    .PARAMETER InputObject
        The data to display in the table. Can be an array of objects or a single object.

    .PARAMETER TableParameters
        A hashtable of parameters to pass to the underlying Table function.
        Supports all parameters of the original Table function.

    .PARAMETER RemoveEmptyColumns
        If specified, columns containing only empty or whitespace values will be removed from the output.

    .PARAMETER CustomColumnWidths
        A hashtable specifying custom widths for specific columns.
        Key: Column name
        Value: Desired width (percentage of total width)
        Remaining width will be distributed among unspecified columns.

    .PARAMETER PassThru
        If specified, returns an object containing metadata about the table operation
        (e.g., which columns were removed).

    .EXAMPLE
        $data | Write-EnhancedTable -TableParameters @{ Name = 'MyTable'; Style = 'Basic' }
        Creates a basic table using default width calculations.

    .EXAMPLE
        $data | Write-EnhancedTable -RemoveEmptyColumns -PassThru
        Creates a table with empty columns removed and returns metadata about removed columns.

    .EXAMPLE
        $data | Write-EnhancedTable -CustomColumnWidths @{ 'Name' = 30; 'Description' = 50 }
        Creates a table with specific column widths, automatically distributing remaining width.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [PSObject]$InputObject,

        [Parameter()]
        [hashtable]$TableParameters = @{},

        [Parameter()]
        [switch]$RemoveEmptyColumns,

        [Parameter()]
        [ValidateScript({
            $_.Values | ForEach-Object {
                if ($_ -isnot [int] -or $_ -lt 1 -or $_ -gt 100) {
                    throw "Column widths must be integers between 1 and 100."
                }
            }
            return $true
        })]
        [hashtable]$CustomColumnWidths = @{},

        [Parameter()]
        [switch]$PassThru
    )

    begin {
        $accumulated = [System.Collections.ArrayList]@()
        $hasProcessed = $false
        $hasCalledTable = $false

        # Get table name for messages
        $tableName = if ($TableParameters.ContainsKey('Name')) {
            $TableParameters.Name
        } else {
            'Unnamed Table'
        }
        Write-PScriboMessage -Message ($localized.ProcessingEnhancedTable -f $tableName)
    }

    process {
        if ($hasProcessed) { return }
        [void]$accumulated.Add($InputObject)
    }

    end {
        if ($hasProcessed) { return }
        $hasProcessed = $true

        # Convert ArrayList to array and handle empty or null input
        $processedObjects = @($accumulated)
        if ($null -eq $processedObjects -or ($processedObjects | Measure-Object).Count -eq 0) {
            Write-PScriboMessage -Message ($localized.NoInputProvided -f $tableName)
            return
        }

        $metadata = @{
            RemovedColumns = @()
            ProcessedObject = $processedObjects
        }

        # Get all unique properties from all objects
        $allProperties = @($processedObjects | ForEach-Object { 
            $obj = $_
            if (($obj.PSObject.Properties | Measure-Object).Count -eq 0) {
                # For objects with no properties, create a default property
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Value' -Value ''
                @('Value')
            } else {
                $obj.PSObject.Properties.Name
            }
        } | Select-Object -Unique)
        
        if (($allProperties | Measure-Object).Count -eq 0) {
            Write-PScriboMessage -Message ($localized.NoPropertiesFound -f $tableName)
            return
        }

        # Ensure all objects have all properties (even if empty)
        $metadata.ProcessedObject = @($processedObjects | ForEach-Object {
            $obj = $_
            $result = [PSCustomObject]@{}
            foreach ($prop in $allProperties) {
                Add-Member -InputObject $result -MemberType NoteProperty -Name $prop -Value $(
                    if ($obj.PSObject.Properties[$prop]) { $obj.$prop } else { '' }
                )
            }
            $result
        })

        # Handle empty columns if requested
        if ($RemoveEmptyColumns) {
            Write-PScriboMessage -Message ($localized.ProcessingEmptyColumns -f $tableName)
            $emptyColumns = @(Find-EmptyTableColumn -InputObject $metadata.ProcessedObject)
            if (($emptyColumns | Measure-Object).Count -gt 0) {
                Write-PScriboMessage -Message ($localized.EmptyColumnsFound -f ($emptyColumns | Measure-Object).Count, $tableName, ($emptyColumns -join ', '))
                $metadata.RemovedColumns = $emptyColumns
                $properties = $metadata.ProcessedObject[0].PSObject.Properties.Name | Where-Object { $_ -notin $emptyColumns }
                $metadata.ProcessedObject = $metadata.ProcessedObject | Select-Object $properties
            }
            else {
                Write-PScriboMessage -Message ($localized.NoEmptyColumnsFound -f $tableName)
            }
        }

        # Calculate column widths if custom widths specified or if columns were removed
        if ($CustomColumnWidths.Count -gt 0 -or ($RemoveEmptyColumns -and ($metadata.RemovedColumns | Measure-Object).Count -gt 0)) {
            Write-PScriboMessage -Message ($localized.ProcessingColumnWidths -f $tableName)
            $TableParameters['ColumnWidths'] = Get-TableColumnWidth -InputObject $metadata.ProcessedObject -CustomWidths $CustomColumnWidths
            Write-PScriboMessage -Message ($localized.ColumnWidthsCalculated -f $tableName, ($TableParameters['ColumnWidths'] -join ', '))
        }

        # Create the table
        if (!$script:hasCalledTable) {
            Write-PScriboMessage -Message "Creating table with $(($metadata.ProcessedObject[0].PSObject.Properties | Measure-Object).Count) columns..."
            $script:hasCalledTable = $true
            $metadata.ProcessedObject | Table @TableParameters
        }

        Write-PScriboMessage -Message ($localized.EnhancedTableCompleted -f $tableName)

        # Return metadata if requested
        if ($PassThru) {
            [PSCustomObject]@{
                RemovedColumns = $metadata.RemovedColumns
                TableParameters = $TableParameters
            }
        }
    }
}
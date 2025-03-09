#requires -Modules PScribo

<#
    Example45 - Write-EnhancedTable automatic column widths and empty column handling

    This example demonstrates the enhanced table functionality including:
    - Automatic column width calculation
    - Empty column removal
    - Combining both features
    - Using metadata about removed columns
#>

Import-Module PScribo

# Define the document
Document 'Example45' {
    # Create a basic style for tables
    TableStyle -Name 'Basic' -HeaderStyle Normal -RowStyle Normal -AlternateRowStyle Normal

    # Example 1: Basic table with automatic column widths
    Paragraph -Style Heading1 'Automatic Column Width Distribution'
    $services = Get-Service | Select-Object -First 5 | Select-Object Name, DisplayName, Status
    Write-EnhancedTable -InputObject $services `
        -CustomColumnWidths @{
            'Name' = 25        # 25% of total width
            'DisplayName' = 55 # 55% of total width
            'Status' = 20      # 20% of total width
        } `
        -TableParameters @{
            Name = 'Services with Custom Widths'
            Style = 'Basic'
        }
    BlankLine

    # Example 2: Table with empty column removal
    Paragraph -Style Heading1 'Empty Column Handling'
    $data = @(
        [PSCustomObject]@{
            Name = 'Item1'
            Description = 'First item'
            EmptyColumn1 = ''
            Value = 100
            EmptyColumn2 = $null
        },
        [PSCustomObject]@{
            Name = 'Item2'
            Description = 'Second item'
            EmptyColumn1 = ''
            Value = 200
            EmptyColumn2 = ''
        }
    )

    # First show which columns are empty
    $emptyColumns = Get-TableEmptyColumn -InputObject $data
    if ($emptyColumns) {
        Paragraph "Note: The following columns contain no data: $($emptyColumns -join ', ')"
    }

    # Create table without empty columns
    Write-EnhancedTable -InputObject $data `
        -RemoveEmptyColumns `
        -TableParameters @{
            Name = 'Data with Empty Columns Removed'
            Style = 'Basic'
        }
    BlankLine

    # Example 3: Combining both features
    Paragraph -Style Heading1 'Combined Features'
    Section 'System Information' {
        # Get some system information that might have empty columns
        $sysInfo = Get-WmiObject Win32_ComputerSystem |
            Select-Object Name, Manufacturer, Model, Description, PrimaryOwnerName |
            Select-Object -First 1

        # Get metadata about what columns were removed
        $result = Write-EnhancedTable -InputObject $sysInfo `
            -RemoveEmptyColumns `
            -CustomColumnWidths @{
                'Name' = 20
                'Manufacturer' = 30
                'Model' = 30
            } `
            -TableParameters @{
                Name = 'System Information'
                Style = 'Basic'
            } `
            -PassThru

        if ($result.RemovedColumns) {
            Paragraph -Style 'Normal' "Note: Some columns were automatically removed as they contained no data: $($result.RemovedColumns -join ', ')"
        }
    }
} | Export-Document -Path $PSScriptRoot -Format Word,Html

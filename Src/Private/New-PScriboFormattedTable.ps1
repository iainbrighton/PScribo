function New-PScriboFormattedTable
{
<#
    .SYNOPSIS
        Creates a formatted table with row and column styling for plugin output/rendering.

    .NOTES
        Maintains backwards compatibility with other plugins that do not require styling/formatting.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Table,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasHeaderRow,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasHeaderColumn
    )
    process
    {
        $tableStyle = $document.TableStyles[$Table.Style]
        return [PSCustomObject] @{
            Id              = $Table.Id
            Name            = $Table.Name
            Number          = $Table.Number
            Type            = $Table.Type
            ColumnWidths    = $Table.ColumnWidths
            Rows            = New-Object -TypeName System.Collections.ArrayList
            Width           = $Table.Width
            Tabs            = if ($null -eq $Table.Tabs) { 0 } else { $Table.Tabs }
            PaddingTop      = $tableStyle.PaddingTop
            PaddingLeft     = $tableStyle.PaddingLeft
            PaddingBottom   = $tableStyle.PaddingBottom;
            PaddingRight    = $tableStyle.PaddingRight
            Align           = $tableStyle.Align
            BorderWidth     = $tableStyle.BorderWidth
            BorderStyle     = $tableStyle.BorderStyle
            BorderColor     = $tableStyle.BorderColor
            Style           = $Table.Style
            HasHeaderRow    = $HasHeaderRow  ## First row has header style applied (list/combo list tables)
            HasHeaderColumn = $HasHeaderColumn  ## First column has header style applied (list/combo list tables)
            HasCaption      = $Table.HasCaption
            Caption         = $Table.Caption
            IsList          = $Table.IsList
            IsKeyedList     = $Table.IsKeyedList
        }
    }
}

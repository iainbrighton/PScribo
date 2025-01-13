function Get-TableColumnWidth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject]
        $InputObject,
        
        [Parameter()]
        [System.Collections.Hashtable]
        $CustomWidths
    )
    process {
        if ($CustomWidths) {
            $totalWidth = ($CustomWidths.Values | Measure-Object -Sum).Sum
            if ($totalWidth -gt 100) {
                throw "Custom widths total $totalWidth exceeds the total available width"
            }
        }
    }
} 
function Get-WordTableConditionalFormattingValue
{
<#
    .SYNOPSIS
        Generates legacy table conditioning formatting value (for LibreOffice).
#>
    [CmdletBinding()]
    [OutputType([System.Int32])]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasFirstRow,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasLastRow,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasFirstColumn,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $HasLastColumn,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoHorizontalBand,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoVerticalBand
    )
    process
    {
        $val = 0
        if ($HasFirstRow)
        {
            $val = $val -bor 0x20
        }
        if ($HasLastRow)
        {
            $val = $val -bor 0x40
        }
        if ($HasFirstColumn)
        {
            $val = $val -bor 0x80
        }
        if ($HasLastColumn)
        {
            $val = $val -bor 0x100
        }
        if ($NoHorizontalBand)
        {
            $val = $val -bor 0x200
        }
        if ($NoVerticalBand)
        {
            $val = $val -bor 0x400
        }
        return $val
    }
}

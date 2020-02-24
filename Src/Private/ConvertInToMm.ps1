function ConvertInToMm
{
<#
    .SYNOPSIS
        Convert inches into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('in')]
        [System.Single] $Inch
    )
    process
    {
        $mm = $Inch * 25.4
        return [System.Math]::Round($mm, 2)
    }
}

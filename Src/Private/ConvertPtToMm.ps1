function ConvertPtToMm
{
<#
    .SYNOPSIS
        Convert points into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('pt')]
        [System.Single] $Point
    )
    process
    {
        return [System.Math]::Round(($Point / 72) * 25.4, 2);
    }
}

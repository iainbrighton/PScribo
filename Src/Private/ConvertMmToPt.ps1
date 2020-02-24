function ConvertMmToPt
{
<#
    .SYNOPSIS
        Convert millimeters into points
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process
    {
        $pt = (ConvertMmToIn $Millimeter) / 0.0138888888888889
        return [System.Math]::Round($pt, 2)
    }
}

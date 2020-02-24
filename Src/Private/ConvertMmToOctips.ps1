function ConvertMmToOctips
{
<#
    .SYNOPSIS
        Convert millimeters into octips

    .NOTES
        1 "octip" = 1/8th pt
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
        $octips = (ConvertMmToIn -Millimeter $Millimeter) * 576
        return [System.Math]::Round($octips, 2)
    }
}

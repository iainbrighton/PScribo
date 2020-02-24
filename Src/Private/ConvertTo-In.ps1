function ConvertTo-In
{
<#
    .SYNOPSIS
        Convert millimeters into inches
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
        $in = $Millimeter / 25.4
        return [System.Math]::Round($in, 2)
    }
}

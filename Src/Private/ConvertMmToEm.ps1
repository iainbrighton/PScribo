function ConvertMmToEm
{
<#
    .SYNOPSIS
        Convert millimeters into em
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
        $em = $Millimeter / 4.23333333333333
        return [System.Math]::Round($em, 2)
    }
}

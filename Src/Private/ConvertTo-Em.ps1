function ConvertTo-Em
{
<#
    .SYNOPSIS
        Convert pixels or millimeters into EMU
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Pixel')]
        [Alias('px')]
        [System.Single] $Pixel,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Millimeter')]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Pixel')
        {
            $em = $pixel * 9525
            return [System.Math]::Round($em, 0)
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Millimeter')
        {
            $em = $Millimeter / 4.23333333333333
            return [System.Math]::Round($em, 2)
        }
    }
}

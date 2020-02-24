function ConvertPxToEm
{
<#
    .SYNOPSIS
        Convert pixels into EMU
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('px')]
        [System.Single] $Pixel
    )
    process
    {
        $em = $pixel * 9525
        return [System.Math]::Round($em, 0)
    }
}

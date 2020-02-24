function ConvertPxToMm
{
<#
    .SYNOPSIS
        Convert pixels into millimeters (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('px')]
        [System.Single] $Pixel,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process
    {
        $mm = (25.4 / $Dpi) * $Pixel
        return [System.Math]::Round($mm, 2);
    }
}

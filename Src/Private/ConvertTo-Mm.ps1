function ConvertTo-Mm
{
<#
    .SYNOPSIS
        Convert points, inches or pixels into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Point')]
        [Alias('pt')]
        [System.Single] $Point,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Inch')]
        [Alias('in')]
        [System.Single] $Inch,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Pixel')]
        [Alias('px')]
        [System.Single] $Pixel,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Pixel')]
        [System.Int16] $Dpi = 96
    )
    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Point')
        {
            return [System.Math]::Round(($Point / 72) * 25.4, 2)
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Inch')
        {
            $mm = $Inch * 25.4
            return [System.Math]::Round($mm, 2)
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Pixel')
        {
            $mm = (25.4 / $Dpi) * $Pixel
            return [System.Math]::Round($mm, 2)
        }
    }
}

function ConvertTo-Twips
{
<#
    .SYNOPSIS
        Convert millimeters into twips

    .NOTES
        1 twip = 1/20th pt
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns','')]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process
    {
        $twips = (ConvertTo-In -Millimeter $Millimeter) * 1440
        return [System.Math]::Round($twips, 2)
    }
}

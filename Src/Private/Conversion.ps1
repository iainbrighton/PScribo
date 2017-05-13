function ConvertPtToMm {
<#
    .SYNOPSIS
        Convert points into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('pt')]
        [System.Single] $Point
    )
    process {

        return [System.Math]::Round(($Point / 72) * 25.4, 2);

    }
} #end function ConvertPtToMm


function ConvertPxToMm {
<#
    .SYNOPSIS
        Convert pixels into millimeters (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('px')]
        [System.Single] $Pixel,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process {

        return [System.Math]::Round((25.4 / $Dpi) * $Pixel, 2);

    }
} #end function ConvertPxToMm


function ConvertInToMm {
<#
    .SYNOPSIS
        Convert inches into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('in')]
        [System.Single] $Inch
    )
    process {

        return [System.Math]::Round($Inch * 25.4, 2);

    }
} #end function ConvertInToMm


function ConvertMmToIn {
<#
    .SYNOPSIS
        Convert millimeters into inches
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {

        return [System.Math]::Round($Millimeter / 25.4, 2);

    }
} #end function ConvertMmToIn


function ConvertMmToPt {
<#
    .SYNOPSIS
        Convert millimeters into points
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {
        return ((ConvertMmToIn $Millimeter) / 0.0138888888888889);

    }
} #end function ConvertMmToPt


function ConvertMmToTwips {
<#
    .SYNOPSIS
        Convert millimeters into twips
    .NOTES
        1 twip = 1/20th pt
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {

        return (ConvertMmToIn -Millimeter $Millimeter) * 1440;

    }
} #end function ConvertMmToTwips


function ConvertMmToOctips {
<#
    .SYNOPSIS
        Convert millimeters into octips
    .NOTES
        1 "octip" = 1/8th pt
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {

        return (ConvertMmToIn -Millimeter $Millimeter) * 576;

    }
} #end function ConvertMmToOctips


function ConvertMmToEm {
<#
    .SYNOPSIS
        Convert millimeters into em
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter
    )
    process {

        return [System.Math]::Round($Millimeter / 4.23333333333333, 2);

    }
} #end function ConvertMmToEm


function ConvertMmToPx {
<#
    .SYNOPSIS
        Convert millimeters into pixels (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Int16])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process {

        $pixels = [System.Int16] ((ConvertMmToIn -Millimeter $Millimeter) * $Dpi);
        if ($pixels -lt 1) { return (1 -as [System.Int16]); }
        else { return $pixels; }

    }
} #end function ConvertMmToPx

function ConvertToInvariantCultureString {
    <#
        .SYNOPSIS
            Convert to a number to a string with a culture-neutral representation #6, #42.
    #>
    [CmdletBinding()]
    param (
        ## The sinle/double
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Object] $Object,

        ## Format string
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Format
    )

    if ($PSBoundParameters.ContainsKey('Format')) {

        return $Object.ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture);
    }
    else {

        return $Object.ToString([System.Globalization.CultureInfo]::InvariantCulture);
    }

} #end function ConvertToInvariantCultureString

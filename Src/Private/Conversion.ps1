function ConvertPtToMm {
<#
    .SYNOPSIS
        Convert points into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('pt')]
        [System.Single] $Point
    )
    process
    {
        return [System.Math]::Round(($Point / 72) * 25.4, 2);
    }
}


function ConvertPxToMm {
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


function ConvertInToMm {
<#
    .SYNOPSIS
        Convert inches into millimeters
#>
    [CmdletBinding()]
    [OutputType([System.Single])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('in')]
        [System.Single] $Inch
    )
    process
    {
        $mm = $Inch * 25.4
        return [System.Math]::Round($mm, 2)
    }
}


function ConvertMmToIn {
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


function ConvertMmToPt {
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


function ConvertMmToTwips {
<#
    .SYNOPSIS
        Convert millimeters into twips

    .NOTES
        1 twip = 1/20th pt
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
        $twips = (ConvertMmToIn -Millimeter $Millimeter) * 1440
        return [System.Math]::Round($twips, 2)
    }
}


function ConvertMmToOctips {
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


function ConvertMmToEm {
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


function ConvertMmToPx {
<#
    .SYNOPSIS
        Convert millimeters into pixels (default 96dpi)
#>
    [CmdletBinding()]
    [OutputType([System.Int16])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('mm','Millimetre')]
        [System.Single] $Millimeter,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int16] $Dpi = 96
    )
    process
    {
        $px = [System.Int16] ((ConvertMmToIn -Millimeter $Millimeter) * $Dpi)
        if ($px -lt 1) { return (1 -as [System.Int16]) }
        else { return $px }

    }
}

function ConvertToInvariantCultureString {
    <#
        .SYNOPSIS
            Convert to a number to a string with a culture-neutral representation #6, #42.
    #>
    [CmdletBinding()]
    param
    (
        ## The sinle/double
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Object] $Object,

        ## Format string
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Format
    )
    process
    {
        if ($PSBoundParameters.ContainsKey('Format'))
        {
            return $Object.ToString($Format, [System.Globalization.CultureInfo]::InvariantCulture);
        }
        else
        {
            return $Object.ToString([System.Globalization.CultureInfo]::InvariantCulture);
        }
    }
}

function ConvertPxToEm {
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

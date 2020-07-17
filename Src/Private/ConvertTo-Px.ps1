function ConvertTo-Px
{
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
        $px = [System.Int16] ((ConvertTo-In -Millimeter $Millimeter) * $Dpi)
        if ($px -lt 1)
        {
            return (1 -as [System.Int16])
        }
        else
        {
            return $px
        }
    }
}

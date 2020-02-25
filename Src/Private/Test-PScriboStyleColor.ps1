function Test-PScriboStyleColor
{
<#
    .SYNOPSIS
        Tests whether a color string is a valid HTML color.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color
    )
    process
    {
        if (Resolve-PScriboStyleColor -Color $Color)
        {
            return $true
        }
        else
        {
            return $false
        }
    }
}

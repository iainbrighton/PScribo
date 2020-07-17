function ConvertTo-InvariantCultureString
{
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

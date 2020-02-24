function Resolve-ImageUri
{
<#
    .SYNOPSIS
        Converts an image path into a Uri.

    .NOTES
        A Uri includes information about whether the path is local etc. This is useful for plugins
        to be able to determine whether to embed images or not.
#>
    [CmdletBinding()]
    [OutputType([System.Uri])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.String] $Path
    )
    process
    {
        if (Test-Path -Path $Path)
        {
            $Path = Resolve-Path -Path $Path
        }
        $uri = New-Object -TypeName System.Uri -ArgumentList @($Path)
        return $uri
    }
}

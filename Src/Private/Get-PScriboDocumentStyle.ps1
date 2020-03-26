function Get-PScriboDocumentStyle
{
<#
    .SYNOPSIS
        Returns document style or table style.

    .NOTES
        Enables testing without having to generate a mock document object!
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Style')]
        [System.String] $Style,

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'TableStyle')]
        [System.String] $TableStyle
    )
    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Style')
        {
            return $Document.Styles[$Style]
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'TableStyle')
        {
            return $Document.TableStyles[$TableStyle]
        }
    }
}

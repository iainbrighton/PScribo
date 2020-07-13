function TOC {
<#
    .SYNOPSIS
        Initializes a new PScribo Table of Contents (TOC) object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name = 'Contents',

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $ClassId = 'TOC'
    )
    process
    {
        Write-PScriboMessage -Message ($localized.ProcessingTOC -f $Name)
        return (New-PScriboTOC @PSBoundParameters)
    }
}

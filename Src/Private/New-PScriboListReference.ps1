function New-PScriboListReference
{
<#
    .SYNOPSIS
        Initializes new PScribo list reference object.

    .DESCRIPTION
        Creates a placeholder reference to a list stored in $pscriboDocument.Lists.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## Display name used in verbose output when processing.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Name,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Number
    )
    process
    {
        $pscriboListReference = [PSCustomObject] @{
            Id                = [System.Guid]::NewGuid().ToString()
            Name              = $Name
            Type              = 'PScribo.ListReference'
            Number            = $Number
        }
        return $pscriboListReference
    }
}

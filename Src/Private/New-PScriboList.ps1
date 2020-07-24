function New-PScriboList
{
<#
    .SYNOPSIS
        Initializes new PScribo list object.

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

        ## List style Name/Id reference.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Numbered list.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Numbered
    )
    process
    {
        $pscriboList = [PSCustomObject] @{
            Id                = [System.Guid]::NewGuid().ToString()
            Name              = $Name
            Type              = 'PScribo.List'
            Style             = $Style
            Items             = (New-Object -TypeName System.Collections.ArrayList)
            IsNumbered        = $Numbered.ToBool()
            IsSectionBreak    = $false
            IsSectionBreakEnd = $false
            IsStyleInherited  = -not $PSBoundParameters.ContainsKey('Style')
            HasStyle          = $PSBoundParameters.ContainsKey('Style')
        }
        return $pscriboList
    }
}

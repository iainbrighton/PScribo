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
        [System.Management.Automation.SwitchParameter] $Numbered,

        ## Numbered list style.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $NumberStyle = $pscriboDocument.DefaultNumberStyle,

        ## Bullet list style.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Circle', 'Dash', 'Disc', 'Square')]
        [System.String] $BulletStyle = 'Disc',

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Level = 1
    )
    process
    {
        $pscriboList = [PSCustomObject] @{
            Id                = [System.Guid]::NewGuid().ToString()
            Name              = $Name
            Type              = 'PScribo.List'
            Items             = (New-Object -TypeName System.Collections.ArrayList)
            Number            = 0
            Level             = $Level
            IsNumbered        = $Numbered.ToBool()
            Style             = $Style
            BulletStyle       = $BulletStyle
            NumberStyle       = $NumberStyle
            IsMultiLevel      = $false
            IsSectionBreak    = $false
            IsSectionBreakEnd = $false
            IsStyleInherited  = -not $PSBoundParameters.ContainsKey('Style')
            HasStyle          = $PSBoundParameters.ContainsKey('Style')
            HasBulletStyle    = -not $Numbered.ToBool()
            HasNumberStyle    = $PSBoundParameters.ContainsKey('NumberStyle')
        }
        return $pscriboList
    }
}

function Section
{
<#
    .SYNOPSIS
        Initializes a new PScribo section object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo section heading/name.
        [Parameter(Mandatory, Position = 0)]
        [System.String] $Name,

        ## PScribo document script block.
        [Parameter(Position = 1)]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),

        ## PScribo style applied to document section.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,

        ## Section is excluded from TOC/section numbering.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ExcludeFromTOC,

        ## Tab indent
        [Parameter()]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0,

        ## Section orientation
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Portrait','Landscape')]
        [System.String] $Orientation
    )
    process
    {
        Write-PScriboMessage -Message ($localized.ProcessingSectionStarted -f $Name)

        $newPScriboSectionParams = @{
            Name        = $Name
            Style       = $Style
            IsExcluded  = $ExcludeFromTOC
            Tabs        = $Tabs
        }
        if ($PSBoundParameters.ContainsKey('Orientation'))
        {
            $newPScriboSectionParams['Orientation'] = $Orientation
        }
        $pscriboSection = New-PScriboSection @newPScriboSectionParams

        foreach ($result in & $ScriptBlock)
        {
            ## Ensure we don't have something errant passed down the pipeline (#29)
            if ($result -is [System.Management.Automation.PSObject])
            {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -match '^PScribo.'))
                {
                    [ref] $null = $pscriboSection.Sections.Add($result)
                }
                else
                {
                    Write-PScriboMessage -Message ($localized.UnexpectedObjectWarning -f $Name) -IsWarning
                }
            }
            else
            {
                Write-PScriboMessage -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), $Name) -IsWarning
            }
        }
        Write-PScriboMessage -Message ($localized.ProcessingSectionCompleted -f $Name)
        return $pscriboSection;
    }
}

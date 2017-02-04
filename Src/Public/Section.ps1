function Section {
<#
    .SYNOPSIS
        Initializes a new PScribo section object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
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
        [System.Management.Automation.SwitchParameter] $ExcludeFromTOC
    )
    begin {

        <#! Section.Internal.ps1 !#>

    } #end begin
    process {

        WriteLog -Message ($localized.ProcessingSectionStarted -f $Name);
        $pscriboSection = New-PScriboSection -Name $Name -Style $Style -IsExcluded:$ExcludeFromTOC;
        foreach ($result in & $ScriptBlock) {

            ## Ensure we don't have something errant passed down the pipeline (#29)
            if ($result -is [System.Management.Automation.PSCustomObject]) {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -match '^PScribo.')) {

                    [ref] $null = $pscriboSection.Sections.Add($result);
                }
                else {

                    WriteLog -Message ($localized.UnexpectedObjectWarning -f $Name) -IsWarning;
                }
            }
            else {

                WriteLog -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), $Name) -IsWarning;
            }
        }
        WriteLog -Message ($localized.ProcessingSectionCompleted -f $Name);
        return $pscriboSection;

    } #end process
} #end function Section

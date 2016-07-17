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
            [ref] $null = $pscriboSection.Sections.Add($result);
        }
        WriteLog -Message ($localized.ProcessingSectionCompleted -f $Name);
        return $pscriboSection;

    } #end process
} #end function Section

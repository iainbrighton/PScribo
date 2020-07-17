function New-PScriboSection
{
<#
    .SYNOPSIS
        Initializes new PScribo section object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo section heading/name.
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## PScribo style applied to document section.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,

        ## Section is excluded from TOC/section numbering.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $IsExcluded,

        ## Tab indent
        [Parameter()]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0,

        ## Section orientation
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Portrait','Landscape')]
        [System.String] $Orientation
    )
    begin
    {
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($PSBoundParameters.ContainsKey('Orientation') -and ($psCallStack[2].FunctionName -ne 'Document<Process>'))
        {
            Write-PScriboMessage -Message $localized.CannotSetOrientationWarning -IsWarning;
            $null = $PSBoundParameters.Remove('Orientation')
        }
    }
    process
    {
        $typeName = 'PScribo.Section';
        $pscriboDocument.Properties['Sections']++;
        $pscriboSection = [PSCustomObject] @{
            Id                = [System.Guid]::NewGuid().ToString();
            Level             = 0;
            Number            = '';
            Name              = $Name;
            Type              = $typeName;
            Style             = $Style;
            Tabs              = $Tabs;
            IsExcluded        = $IsExcluded;
            Sections          = (New-Object -TypeName System.Collections.ArrayList);
            Orientation       = if ($PSBoundParameters.ContainsKey('Orientation')) { $Orientation } else { $script:currentOrientation }
            IsSectionBreak    = $false;
            IsSectionBreakEnd = $false;
        }

        ## Has the orientation changed from the parent scope
        if ($PSBoundParameters.ContainsKey('Orientation') -and ($Orientation -ne $script:currentOrientation))
        {
            $pscriboSection.IsSectionBreak = $true;
            $pscriboDocument.Properties['Pages']++;
            $script:currentOrientation = $Orientation;
        }
        return $pscriboSection;
    }
}

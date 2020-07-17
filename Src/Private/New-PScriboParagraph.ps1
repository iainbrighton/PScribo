function New-PScriboParagraph
{
<#
    .SYNOPSIS
        Initializes a new PScribo paragraph object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','ScriptBlock')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo paragraph run script block.
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,

        ## Paragraph Id.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString(),

        ## Paragraph style Name/Id reference.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,

        ## Tab indent
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoIncrementCounter
    )
    process
    {
        if (-not $NoIncrementCounter)
        {
            $pscriboDocument.Properties['Paragraphs']++
        }

        $pscriboParagraph = [PSCustomObject] @{
            Id                = $Id
            Type              = 'PScribo.Paragraph'
            Style             = $Style
            Tabs              = $Tabs
            Sections          = (New-Object -TypeName System.Collections.ArrayList)
            Orientation       = $script:currentOrientation
            IsSectionBreakEnd = $false
        }
        return $pscriboParagraph
    }
}

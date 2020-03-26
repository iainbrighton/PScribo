function Get-PScriboHeaderFooter
{
<#
    .SYNOPSIS
        Returns the specified PScribo header/footer object.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'DefaultHeader')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageHeader')]
        [System.Management.Automation.SwitchParameter] $Header,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'DefaultFooter')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageFooter')]
        [System.Management.Automation.SwitchParameter] $Footer,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageHeader')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageFooter')]
        [System.Management.Automation.SwitchParameter] $FirstPage
    )
    process
    {
        if ($FirstPage)
        {
            if ($Header -and ($Document.Header.HasFirstPageHeader))
            {
                return $Document.Header.FirstPageHeader
            }
            elseif ($Footer -and ($Document.Footer.HasFirstPageFooter))
            {
                return $Document.Footer.FirstPageFooter
            }
        }
        else
        {
            if ($Header -and ($Document.Header.HasDefaultHeader))
            {
                return $Document.Header.DefaultHeader
            }
            elseif ($Footer -and ($Document.Footer.HasDefaultFooter))
            {
                return $Document.Footer.DefaultFooter
            }
        }
    }
}

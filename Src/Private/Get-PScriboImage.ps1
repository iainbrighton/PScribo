function Get-PScriboImage
{
<#
    .SYNOPSIS
        Retrieves PScribo.Images in a document/section
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.PSObject[]] $Section,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $Id
    )
    process
    {
        foreach ($subSection in $Section)
        {
            if ($subSection.Type -eq 'PScribo.Image')
            {
                if ($PSBoundParameters.ContainsKey('Id'))
                {
                    if ($subSection.Id -in $Id)
                    {
                        Write-Output -InputObject $subSection
                    }
                }
                else
                {
                    Write-Output -InputObject $subSection
                }
            }
            elseif ($subSection.Type -eq 'PScribo.Section')
            {
                if ($subSection.Sections.Count -gt 0)
                {
                    ## Recursively search subsections
                    $PSBoundParameters['Section'] = $subSection.Sections
                    Get-PScriboImage @PSBoundParameters
                }
            }
        }
    }
}

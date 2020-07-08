function Footer
{
<#
    .SYNOPSIS
        Initializes a new PScribo footer object.

    .NOTES
        Footer sections can only contain Paragraph and Table elements.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo document script block.
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),

        ## Set the default page footer.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.Management.Automation.SwitchParameter] $Default,

        ## Set the First page footer.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FirstPage')]
        [System.Management.Automation.SwitchParameter] $FirstPage,

        ## PScribo inserts a default blank line between the footer and main docoument.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoSpace,

        ## Include default footer on the first page.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.Management.Automation.SwitchParameter] $IncludeOnFirstPage
    )
    process
    {
        Write-PScriboMessage -Message $localized.ProcessingFooterStarted

        $pscriboFooter = New-PScriboHeaderFooter -Footer

        if (-not $NoSpace)
        {
            ## Add a blank line before the footer text
            [ref] $null = $pscriboFooter.Sections.Add((New-PScriboBlankLine))
        }

        foreach ($result in & $ScriptBlock)
        {
            ## Headers/footers only support paragraphs and tables
            if ($result -is [System.Management.Automation.PSObject])
            {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -in 'PScribo.Paragraph','PScribo.Table'))
                {
                    [ref] $null = $pscriboFooter.Sections.Add($result)
                }
                else
                {
                    Write-PScriboMessage -Message ($localized.UnexpectedObjectWarning -f 'Footer') -IsWarning
                }
            }
            else
            {
                Write-PScriboMessage -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), 'Footer') -IsWarning
            }
        }

        if ($FirstPage -or $IncludeOnFirstPage)
        {
            if ($pscriboDocument.Footer.HasFirstPageFooter -eq $true)
            {
                Write-PScriboMessage $localized.FirstPageFooterOverwriteWarning -IsWarning
            }
            $pscriboDocument.Footer.HasFirstPageFooter = $true
            $pscriboDocument.Footer.FirstPageFooter = $pscriboFooter
        }

        if ($Default)
        {
            if ($pscriboDocument.Footer.HasDefaultFooter -eq $true)
            {
                Write-PScriboMessage $localized.DefaultFooterOverwriteWarning -IsWarning
            }
            $pscriboDocument.Footer.HasDefaultFooter = $true
            $pscriboDocument.Footer.DefaultFooter = $pscriboFooter
        }

        Write-PScriboMessage -Message $localized.ProcessingFooterCompleted
    }
}

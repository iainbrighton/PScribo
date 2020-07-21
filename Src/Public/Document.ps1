function Document
{
<#
    .SYNOPSIS
        Initializes a new PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo document name
        [Parameter(Mandatory, Position = 0)]
        [System.String] $Name,

        ## PScribo document DSL script block containing Section, Paragraph and/or Table etc. commands.
        [Parameter(Position = 1)]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),

        ## PScribo document Id
        [Parameter()]
        [System.String] $Id = $Name.Replace(' ','')
    )
    process
    {
        $pluginName = 'Document'
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        $pscriboDocument = New-PScriboDocument -Name $Name -Id $Id
        $pscriboDocument.Properties['Pages']++
        $script:currentOrientation = $pscriboDocument.Options['PageOrientation']

        ## Call the Document script block
        foreach ($result in & $ScriptBlock)
        {
            ## Ensure we don't have something errant passed down the pipeline (#29)
            if ($result -is [System.Management.Automation.PSObject])
            {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -match '^PScribo.'))
                    {
                        [ref] $null = $pscriboDocument.Sections.Add($result)
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

        Invoke-PScriboSection

        ## Process IsSectionBreakEnd (for Word plugin)
        if ($pscriboDocument.Sections.Count -gt 0)
        {
            $previousPScriboSection = $pscriboDocument.Sections[0]
            for ($i = 0; $i -lt $pscriboDocument.Sections.Count; $i++)
            {
                $pscriboSection = $pscriboDocument.Sections[$i]
                if ($pscriboSection.Type -in 'PScribo.Section','PScribo.Paragraph','PScriboTable')
                {
                    if (($null -ne $pscriboSection.PSObject.Properties['IsSectionBreak']) -and ($pscriboSection.IsSectionBreak))
                    {
                        if (($previousPScriboSection.Type -in 'PScribo.Paragraph','PScribo.Table') -or
                            ($previousPScriboSection.Sections.Count -eq 0))
                        {
                            ## Set the last childless section, paragraph or table as the section end
                            $previousPScriboSection.IsSectionBreakEnd = $true
                        }
                        else
                        {
                            ## Set the last child section/paragraph element as the section end
                            Set-PScriboSectionBreakEnd -Section $previousPScriboSection
                        }
                    }
                    $previousPScriboSection = $pscriboSection
                }
            }
        }

        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $pscriboDocument.Name)
        $stopwatch.Stop()

        if ($stopwatch.Elapsed.TotalSeconds -gt 90)
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeMinutes -f $stopwatch.Elapsed.TotalMinutes)
        }
        else
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeSeconds -f $stopwatch.Elapsed.TotalSeconds)
        }

        return $pscriboDocument
    }
}

function Header
{
<#
    .SYNOPSIS
        Initializes a new PScribo header object.

    .NOTES
        Header sections can only contain Paragraph and Table elements.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo document script block.
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),

        ## Set the default page header.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.Management.Automation.SwitchParameter] $Default,

        ## Set the First page header.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FirstPage')]
        [System.Management.Automation.SwitchParameter] $FirstPage,

        ## PScribo inserts a default blank line between the header and main docoument.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoSpace,

        ## Include default header on the first page.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.Management.Automation.SwitchParameter] $IncludeOnFirstPage
    )
    process
    {
        Write-PScriboMessage -Message $localized.ProcessingHeaderStarted

        $pscriboHeader = New-PScriboHeaderFooter -Header

        foreach ($result in & $ScriptBlock)
        {
            ## Headers/footers only support paragraphs and tables
            if ($result -is [System.Management.Automation.PSObject])
            {
                if (('Id' -in $result.PSObject.Properties.Name) -and
                    ('Type' -in $result.PSObject.Properties.Name) -and
                    ($result.Type -in 'PScribo.Paragraph','PScribo.Table'))
                {
                    [ref] $null = $pscriboHeader.Sections.Add($result)
                }
                else
                {
                    Write-PScriboMessage -Message ($localized.UnexpectedObjectWarning -f 'Header') -IsWarning
                }
            }
            else
            {
                Write-PScriboMessage -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), 'Header') -IsWarning
            }
        }

        if (-not $NoSpace)
        {
            ## Add a blank line after the header text
            [ref] $null = $pscriboHeader.Sections.Add((New-PScriboBlankLine))
        }

        if ($FirstPage -or $IncludeOnFirstPage)
        {
            if ($pscriboDocument.Header.HasFirstPageHeader -eq $true)
            {
                Write-PScriboMessage $localized.FirstPageHeaderOverwriteWarning -IsWarning
            }
            $pscriboDocument.Header.HasFirstPageHeader = $true
            $pscriboDocument.Header.FirstPageHeader = $pscriboHeader
        }

        if ($Default)
        {
            if ($pscriboDocument.Header.HasDefaultHeader -eq $true)
            {
                Write-PScriboMessage $localized.DefaultHeaderOverwriteWarning -IsWarning
            }
            $pscriboDocument.Header.HasDefaultHeader = $true
            $pscriboDocument.Header.DefaultHeader = $pscriboHeader
        }

        Write-PScriboMessage -Message $localized.ProcessingHeaderCompleted
    }
}

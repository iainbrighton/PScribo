function Paragraph {
<#
    .SYNOPSIS
        Initializes a new PScribo paragraph object.
#>
    [CmdletBinding(DefaultParameterSetName = 'Legacy')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo paragraph run script block.
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'ParagraphRun')]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,

        ## Paragraph Id and Xml element name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Legacy')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## Paragraph text. If empty $Name/Id will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'Legacy')]
        [AllowNull()]
        [System.String] $Text = $null,

        ## Output value override, i.e. for Xml elements. If empty $Text will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 2, ParameterSetName = 'Legacy')]
        [AllowNull()]
        [System.String] $Value = $null,

        ## Paragraph style Name/Id reference.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ParagraphRun')]
        [AllowNull()]
        [System.String] $Style = $null,

        ## DEPRECATED - Use paragraph runs (Text)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [System.Management.Automation.SwitchParameter] $NoNewLine,

        ## Override the bold style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [System.Management.Automation.SwitchParameter] $Bold,

        ## Override the italic style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [System.Management.Automation.SwitchParameter] $Italic,

        ## Override the underline style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [System.Management.Automation.SwitchParameter] $Underline,

        ## Override the font name(s)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Font,

        ## Override the font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [AllowNull()]
        [System.UInt16] $Size = $null,

        ## Override the font color/colour
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [AllowNull()]
        [System.String] $Color = $null,

        ## Tab indent
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Legacy')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ParagraphRun')]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0
    )
    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Legacy')
        {
            $paragraphId = $Name
            $paragraphDisplayName = $Name
            if ($Name.Length -gt 40)
            {
                $paragraphId = $Name.Substring(0,40)
                $paragraphDisplayName = '{0}[..]' -f$Name.Substring(0,36)
            }
            Write-PScriboMessage -Message ($localized.ProcessingParagraph -f $paragraphDisplayName)

            if ($NoNewLine -eq $true)
            {
                Write-PScriboMessage -Message $localized.NoNewLineDeprecatedWarning -IsWarning
            }
            if ($PSBoundParameters.ContainsKey('Value'))
            {
                Write-PScriboMessage -Message $localized.ValueParameterRemovedWarning -IsWarning
            }

            ## Create paragraph
            $newPScriboParagraphParams = @{
                Id          = $paragraphId
                ScriptBlock = { }
                Tabs        = $Tabs
            }
            if ($PSBoundParameters.ContainsKey('Style'))
            {
                $newPScriboParagraphParams['Style'] = $Style
            }
            $paragraph = New-PScriboParagraph @newPScriboParagraphParams

            ## Create a single run
            $paragraphRunText = $Name
            if ($PSBoundParameters.ContainsKey('Text'))
            {
                $paragraphRunText = $Text
            }
            $null = $PSBoundParameters.Remove('Name')
            $null = $PSBoundParameters.Remove('Text')
            $null = $PSBoundParameters.Remove('Value')
            $null = $PSBoundParameters.Remove('Style')
            $null = $PSBoundParameters.Remove('Tabs')
            $null = $PSBoundParameters.Remove('NoNewLine')
            $paragraphRun = New-PScriboParagraphRun -Text $paragraphRunText @PSBoundParameters
            $paragraphRun.IsParagraphRunEnd = $true
            $null = $paragraph.Sections.Add($paragraphRun)

            return $paragraph
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'ParagraphRun')
        {
            Write-PScriboMessage -Message ($localized.ProcessingParagraphRunsStarted)
            $paragraph = New-PScriboParagraph @PSBoundParameters

            foreach ($result in & $ScriptBlock)
            {
                ## Ensure we don't have something errant passed down the pipeline
                if ($result -is [System.Management.Automation.PSObject])
                {
                    if (('Type' -in $result.PSObject.Properties.Name) -and
                        ($result.Type -eq 'PScribo.ParagraphRun'))
                    {
                        [ref] $null = $paragraph.Sections.Add($result)
                    }
                    elseif (('Type' -in $result.PSObject.Properties.Name) -and
                        ($result.Type -match '^PScribo\.'))
                    {
                        Write-PScriboMessage -Message ($localized.UnsupportedPScriboTypeWarning -f $result.Type, 'Paragraph') -IsWarning
                    }
                    else
                    {
                        Write-PScriboMessage -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), 'Paragraph') -IsWarning
                    }
                }
                else
                {
                    Write-PScriboMessage -Message ($localized.UnexpectedObjectTypeWarning -f $result.GetType(), 'Paragraph') -IsWarning
                }
            }

            if ($paragraph.Sections.Count -gt 0)
            {
                ## Set the paragraph Id to the text of the first run
                if ($paragraph.Sections[0].Text.Length -gt 40)
                {
                    $paragraphDisplayName = '{0}[..]' -f $paragraph.Sections[0].Text.Substring(0,36)
                }
                else
                {
                    $paragraphDisplayName = $paragraph.Sections[0].Text
                }
                $paragraph.Id = $paragraphDisplayName
                $paragraph.Sections[-1].IsParagraphRunEnd = $true
            }

            Write-PScriboMessage -Message ($localized.ProcessingParagraphRunsCompleted)
            return $paragraph
        }
        Write-PScriboMessage -Message ($localized.ProcessingParagraph -f $paragraphDisplayName)

        return (New-PScriboParagraph @PSBoundParameters)
    }
}

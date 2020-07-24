function List
{
<#
    .SYNOPSIS
        Initializes a new PScribo bulleted or numbered List object.
#>
    [CmdletBinding(DefaultParameterSetName = 'Item')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## List item(s).
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Item')]
        [ValidateNotNull()]
        [System.String[]] $Item,

        ## PScribo nested list/items.
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,

        ## Display name used in verbose output when processing.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Name,

        ## List style Name/Id reference.
        [Parameter(Position = 1, ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Numbered list.
        [Parameter(Position = 2, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Numbered
    )
    begin
    {
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($psCallStack[1].FunctionName -notin 'List<Process>','Document<Process>','Section<Process>')
        {
            throw $localized.ListRootError
        }
    }
    process
    {
        $null = $PSBoundParameters.Remove('ScriptBlock')
        $null = $PSBoundParameters.Remove('Item')
        $pscriboList = New-PScriboList @PSBoundParameters

        if ($PSCmdlet.ParameterSetName -eq 'Item')
        {
            foreach ($listItem in $Item)
            {
                $pscriboListItem = New-PScriboItem -Text $listItem
                [ref] $null = $pscriboList.Items.Add($pscriboListItem)
            }
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'List')
        {
            foreach ($result in & $ScriptBlock)
            {
                ## Ensure we don't have something errant passed down the pipeline (#29)
                if ($result -is [System.Management.Automation.PSObject])
                {
                    if (('Id' -in $result.PSObject.Properties.Name) -and
                        ('Type' -in $result.PSObject.Properties.Name) -and
                        ($result.Type -match '^PScribo.'))
                    {
                        [ref] $null = $pscriboList.Items.Add($result)
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
        }

        ## Only process list levels on the root list object
        if ($psCallStack[1].FunctionName -in 'Document<Process>','Section<Process>')
        {
            Write-PScriboMessage -Message ($localized.ProcessingList -f $pscriboList.Name)
            $hasItem = $false
            $itemNumber = 0

            foreach ($listItem in $pscriboList.Items)
            {
                if ($listItem.Type -eq 'PScribo.Item')
                {
                    $itemNumber++
                    $listItem.Level = 1
                    $listItem.Index = $itemNumber
                    $listItem.Number = $itemNumber.ToString()

                    $hasItem = $true
                }
                elseif ($listItem.Type -eq 'PScribo.List')
                {
                    if ($hasItem)
                    {
                        Invoke-PScriboListLevel -List $listItem -Number $itemNumber.ToString()
                    }
                    else
                    {
                        Write-PScriboMessage -Message $localized.NoPriorListItemWarning -IsWarning
                    }
                }
            }
        }

        return $pscriboList
    }
}

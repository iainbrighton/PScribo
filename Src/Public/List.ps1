function List
{
<#
    .SYNOPSIS
        Initializes a new PScribo bulleted or numbered List object.
#>
    [CmdletBinding(DefaultParameterSetName = 'BulletItem')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## List item(s).
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ItemBullet')]
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ItemNumbered')]
        [ValidateNotNull()]
        [System.String[]] $Item,

        ## PScribo nested list/items.
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'ListBullet')]
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'ListNumbered')]
        [ValidateNotNull()]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,

        ## Display name used in verbose output when processing.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Name,

        ## List style Name/Id reference.
        [Parameter(Position = 1, ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Numbered list.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ItemNumbered')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ListNumbered')]
        [System.Management.Automation.SwitchParameter] $Numbered,

        ## Numbered list style.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ItemNumbered')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ListNumbered')]
        [System.String] $NumberStyle,

        ## Bullet list style.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ItemBullet')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ListBullet')]
        [ValidateSet('Circle', 'Dash', 'Disc', 'Square')]
        [System.String] $BulletStyle = 'Disc'
    )
    begin
    {
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($psCallStack[1].FunctionName -notin 'List<Process>','Document<Process>','Section<Process>')
        {
            Write-Warning $psCallStack[1].FunctionName
            throw $localized.ListRootError
        }
    }
    process
    {
        $null = $PSBoundParameters.Remove('ScriptBlock')
        $null = $PSBoundParameters.Remove('Item')

        $pscriboList = New-PScriboList @PSBoundParameters

        if ($PSCmdlet.ParameterSetName -in 'ItemBullet','ItemNumbered')
        {
            foreach ($listItem in $Item)
            {
                $pscriboListItem = New-PScriboItem -Text $listItem
                [ref] $null = $pscriboList.Items.Add($pscriboListItem)
            }
        }
        elseif ($PSCmdlet.ParameterSetName -in 'ListBullet','ListNumbered')
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
                        $pscriboList.IsMultiLevel = $true
                        Invoke-PScriboListLevel -List $listItem -Number $itemNumber.ToString()
                    }
                    else
                    {
                        Write-PScriboMessage -Message $localized.NoPriorListItemWarning -IsWarning
                    }
                }
            }

            ## Store lists for Word numbering.xml
            $pscriboDocument.Properties['Lists']++
            $pscriboList.Number = $pscriboDocument.Properties['Lists']
            [ref] $null = $pscriboDocument.Lists.Add($pscriboList)

            ## Return list reference
            $pscriboListReference = New-PScriboListReference -Name $pscriboList.Name -Number $pscriboList.Number
            return $pscriboListReference
        }
        else
        {
            return $pscriboList
        }
    }
}

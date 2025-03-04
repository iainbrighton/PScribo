function Invoke-PScriboListLevel
{
<#
    .SYNOPSIS
        Nested function that processes each nested list numbering.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $List,

        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [System.String] $Number
    )
    process
    {
        $level = $Number.Split('.').Count +1
        $processingListPadding = ''.PadRight($level -1, ' ')
        $processingListMessage = $localized.ProcessingList -f $List.Name
        Write-PScriboMessage -Message ('{0}{1}' -f $processingListPadding, $processingListMessage)

        $itemNumber = 0
        $numberString = $Number
        $hasItem = $false

        foreach ($item in $List.Items)
        {
            if ($item.Type -eq 'PScribo.Item')
            {
                $itemNumber++
                $item.Level = $level
                $item.Index = $itemNumber
                $item.Number = ('{0}.{1}' -f $Number, $itemNumber).TrimStart('.')

                $numberString = $item.Number
                $hasItem = $true
            }
            elseif ($item.Type -eq 'PScribo.List')
            {
                if ($hasItem)
                {
                    Invoke-PScriboListLevel -List $item -Number $numberString
                }
                else
                {
                    Write-PScriboMessage -Message $localized.NoPriorListItemWarning -IsWarning
                }
            }
        }
    }
}

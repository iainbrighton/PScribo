function Out-TextList
{
<#
    .SYNOPSIS
        Output formatted text list.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','NumberStyle')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','Options')]
    param
    (
        ## Section to output
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Indent = 2,

        ## Number style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $NumberStyle,

        ## Bullet style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $BulletStyle
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboTextOption
        }
    }
    process
    {
        $listBuilder = New-Object -TypeName System.Text.StringBuilder
        $leader = ''.PadRight($Indent, ' ')

        if ($List.IsNumbered)
        {
            if ($List.HasNumberStyle)
            {
                $NumberStyle = $List.NumberStyle
            }
            elseif (-not $PSBoundParameters.ContainsKey('Style'))
            {
                $NumberStyle = $Document.DefaultNumberStyle
            }
            $style = $Document.NumberStyles[$NumberStyle]

            $maxItemNumberLength = Get-PScriboListItemMaximumLength -List $List -NumberStyle $style

            $outTextListParams = @{
                NumberStyle = $NumberStyle
            }
        }
        else
        {
            if ($List.HasBulletStyle)
            {
                $BulletStyle = $List.BulletStyle
            }

            switch ($BulletStyle)
            {
                Circle
                {
                    $numberString = 'o'
                    break
                }
                Dash
                {
                    $numberString = '-'
                    break
                }
                Disc
                {
                    $numberString = '*'
                    break
                }
                Default
                {
                    ## Square style is not supported in Text so default to the browser's rendering engine
                    $numberString = '*'
                }
            }

            $outTextListParams = @{
                BulletStyle = $BulletStyle
            }
        }

        foreach ($item in $List.Items)
        {
            if ($item.Type -eq 'PScribo.Item')
            {
                if ($List.IsNumbered)
                {
                    $padding = ''
                    $itemNumber = ConvertFrom-NumberStyle -Value $item.Index -NumberStyle $style
                    $paddingLength = ($maxItemNumberLength) - $itemNumber.Length
                    if ($paddingLength -gt 0)
                    {
                        $padding = ''.PadRight($paddingLength, ' ')
                    }

                    if ($style.Align -eq 'Left')
                    {
                        $numberString = '{0}{1}' -f $itemNumber, $padding
                    }
                    elseif ($style.Align -eq 'Right')
                    {
                        $numberString = '{0}{1}' -f $padding, $itemNumber
                    }
                }

                [ref] $null = $listBuilder.AppendFormat('{0}{1} {2}', $leader, $numberString, $item.Text).AppendLine()
            }
            else
            {
                $newIndent = $Indent + 2
                if ($List.IsNumbered)
                {
                    $newIndent = $Indent + $maxItemNumberLength +1
                }
                $nestedList = Out-TextList -List $item -Indent $newIndent @outTextListParams
                [ref] $null = $listBuilder.Append($nestedList)
            }
        }
        [ref] $null = $listBuilder.AppendLine()
        return $listBuilder.ToString()
    }
}

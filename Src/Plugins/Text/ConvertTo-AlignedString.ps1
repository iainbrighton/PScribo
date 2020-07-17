function ConvertTo-AlignedString
{
<#
    .SYNOPSIS
        Justifies and indents a block of text using the specified alignment properties.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Object[]] $InputObject,

        ## Tab indents
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0,

        ## Tab size
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $TabSize = 4,

        ## Text alignment
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right','Justify')]
        [string] $Align = 'Left',

        ## Text width
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Width = ($Host.UI.RawUI.BufferSize.Width -1),

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoNewLine
    )
    begin
    {
        if ($PSBoundParameters.ContainsKey('Debug'))
        {
            $DebugPreference = 'Continue'
        }
    }
    process
    {
        $tabSpaces = ''.PadRight(($Tabs * $TabSize), ' ')
        $usableWidth = $Width - ($Tabs * $TabSize)

        [System.Text.StringBuilder] $textBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [System.Text.StringBuilder] $lineBuilder = New-Object -TypeName 'System.Text.StringBuilder'

        foreach ($line in ($InputObject -split '\r\n?|\n'))
        {
            $currentPosition = 0
            [ref] $null = $lineBuilder.Clear()

            $convertToJustifiedStringParams = @{
                Align       = $Align
                Width       = $usableWidth
            }

            ## Break on spaces and then at maximum "usable" length
            [System.String[]] $words = $line -split '\s' -split "(.{$usableWidth})" |
                        Where-Object { -not [System.String]::IsNullOrEmpty($_) }

            for ($w = 0; $w -lt $words.Count; $w++)
            {
                $word = $words[$w]
                $isLastWord = $w -eq ($words.Count -1)

                Write-Debug "Word: '$word', Length: $($word.Length), IsLastWord: $isLastWord"
                $newPosition = $currentPosition + $word.Length

                if ($currentPosition -gt 0)
                {
                    ## We need to prefix the word with a space if we're not at the start of the line
                    $newPosition += 1
                }

                if ($newPosition -lt $usableWidth)
                {
                    ## We haven't reached the end of the line, so just append
                    if ($currentPosition -gt 0)
                    {
                        ## We're not at the start of the line, so prefix word with a space
                        [ref] $null = $lineBuilder.Append(' ')
                    }

                    [ref] $null = $lineBuilder.Append($word)

                    if ($isLastWord)
                    {
                        ## Output what we have thus far
                        $convertToJustifiedStringParams['InputObject'] = $lineBuilder.ToString()
                        $justifiedString = ConvertTo-JustifiedString @convertToJustifiedStringParams
                        [ref] $null = $textBuilder.Append($tabSpaces).Append($justifiedString)

                        if (-not $NoNewLine)
                        {
                            [ref] $null = $textBuilder.AppendLine()
                        }
                    }

                    $currentPosition = $newPosition
                }
                elseif ($newPosition -eq $usableWidth)
                {
                    ## We're bang on the end of the line, therefore can't justify
                    if ($currentPosition -gt 0)
                    {
                        [ref] $null = $lineBuilder.Append(' ')
                    }

                    [ref] $null = $lineBuilder.Append($word)
                    [ref] $null = $textBuilder.Append($tabSpaces).Append($lineBuilder.ToString())

                    if ($isLastWord -and (-not $NoNewLine))
                    {
                        [ref] $null = $textBuilder.AppendLine()
                    }
                    elseif (-not $isLastWord)
                    {
                        [ref] $null = $textBuilder.AppendLine()
                    }

                    [ref] $null = $lineBuilder.Clear()
                    $currentPosition = 0
                }
                else
                {
                    ## We're over the end of the line, so justify and start a new line
                    $convertToJustifiedStringParams['InputObject'] = $lineBuilder.ToString()
                    $justifiedString = ConvertTo-JustifiedString @convertToJustifiedStringParams
                    [ref] $null = $textBuilder.Append($tabSpaces).Append($justifiedString)

                    if ($isLastWord)
                    {
                        ## Just a single word on a new line
                        $convertToJustifiedStringParams['InputObject'] = $word
                        $justifiedString = ConvertTo-JustifiedString @convertToJustifiedStringParams
                        [ref] $null = $textBuilder.AppendLine().Append($tabSpaces).Append($justifiedString)

                        if (-not $NoNewLine)
                        {
                            [ref] $null = $textBuilder.AppendLine()
                        }
                    }
                    else
                    {
                        ## Start a new line
                        [ref] $null = $textBuilder.AppendLine()
                        [ref] $null = $lineBuilder.Clear().Append($word)
                        $currentPosition = $word.Length
                    }
                }
            }
        }
        return $textBuilder.ToString()
    }
}

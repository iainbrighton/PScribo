function Out-MarkdownLineBreak
{
<#
    .SYNOPSIS
        Output formatted line break text.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param ( )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboMarkdownOption
        }
    }
    process
    {
        $lineBreakSeparatorWidth = $options.LineBreakSeparatorWidth

        if ($options.LineBreakSeparatorWidth -gt $options.TextWidth)
        {
            $lineBreakSeparatorWidth = $options.TextWidth
        }
        $linebreakBuilder = New-Object -TypeName System.Text.StringBuilder
        $lineBreak = ''.PadRight($lineBreakSeparatorWidth, $options.LineBreakSeparator)
        [ref] $null = $linebreakBuilder.Append($lineBreak).AppendLine().AppendLine()

        $script:currentPScriboObject = 'PScribo.LineBreak'
        return $linebreakBuilder.ToString()
    }
}

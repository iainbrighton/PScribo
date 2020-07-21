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
        $linebreakBuilder = New-Object -TypeName System.Text.StringBuilder
        $lineBreak = ''.PadRight($options.LineBreakSeparatorWidth, $options.LineBreakSeparator)
        [ref] $null = $linebreakBuilder.Append($lineBreak).AppendLine().AppendLine()
        return $linebreakBuilder.ToString()
    }
}

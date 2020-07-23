function Out-MarkdownPageBreak
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
        $pagebreakBuilder = New-Object -TypeName System.Text.StringBuilder

        $script:currentPageNumber++
        $pageBreakSeparatorWidth = $options.PageBreakSeparatorWidth
        if ($options.PageBreakSeparatorWidth -gt $options.TextWidth)
        {
            $pageBreakSeparatorWidth = $options.TextWidth
        }
        $pageBreak = ''.PadRight($pageBreakSeparatorWidth, $options.PageBreakSeparator)
        [ref] $null = $pagebreakBuilder.Append($pageBreak).AppendLine().AppendLine()

        $script:currentPScriboObject = 'PScribo.PageBreak'
        return $pageBreakBuilder.ToString()
    }
}

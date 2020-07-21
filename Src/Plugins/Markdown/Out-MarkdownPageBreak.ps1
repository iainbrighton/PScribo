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

        #$isFirstPage = $currentPageNumber -eq 1
        #$pageFooter = Out-TextHeaderFooter -Footer -FirstPage:$isFirstPage
        # [ref] $null = $pageBreakBuilder.Append($pageFooter)

        $script:currentPageNumber++
        $pageBreak = ''.PadRight($options.PageBreakSeparatorWidth, $options.PageBreakSeparator)
        [ref] $null = $pagebreakBuilder.Append($pageBreak).AppendLine().AppendLine()

        # $pageHeader = Out-TextHeaderFooter -Header
        # [ref] $null = $pageBreakBuilder.Append($pageHeader)

        return $pageBreakBuilder.ToString()
    }
}

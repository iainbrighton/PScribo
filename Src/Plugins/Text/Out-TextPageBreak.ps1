function Out-TextPageBreak
{
<#
    .SYNOPSIS
        Output formatted line break text.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param ( )
    process
    {
        $pagebreakBuilder = New-Object -TypeName System.Text.StringBuilder

        $isFirstPage = $currentPageNumber -eq 1
        $pageFooter = Out-TextHeaderFooter -Footer -FirstPage:$isFirstPage
        [ref] $null = $pageBreakBuilder.Append($pageFooter)

        $script:currentPageNumber++
        [ref] $null = $pagebreakBuilder.Append((Out-TextLineBreak)).AppendLine()

        $pageHeader = Out-TextHeaderFooter -Header
        [ref] $null = $pageBreakBuilder.Append($pageHeader)

        return $pageBreakBuilder.ToString()
    }
}

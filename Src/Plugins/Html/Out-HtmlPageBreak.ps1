function Out-HtmlPageBreak
{
<#
    .SYNOPSIS
        Output formatted Html page break.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.String] $Orientation
    )
    process
    {
        [System.Text.StringBuilder] $pageBreakBuilder = New-Object 'System.Text.StringBuilder'

        [ref] $null = $pageBreakBuilder.Append('</div>') ## Canvas
        $isFirstPage = $currentPageNumber -eq 1
        [ref] $null = $pageBreakBuilder.Append((Out-HtmlHeaderFooter -Footer -FirstPage:$isFirstPage))

        $script:currentPageNumber++
        [ref] $null = $pageBreakBuilder.Append('</div></div>')
        $topMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $Document.Options['MarginTop']) -Format 'f2'
        $leftMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $Document.Options['MarginLeft']) -Format 'f2'
        $bottomMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $Document.Options['MarginBottom']) -Format 'f2'
        $rightMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $Document.Options['MarginRight']) -Format 'f2'
        [ref] $null = $pageBreakBuilder.AppendFormat('<div class="{0}">', $Orientation.ToLower())
        [ref] $null = $pageBreakBuilder.AppendFormat('<div class="{0}" style="padding-top: {1}rem; padding-left: {2}rem; padding-bottom: {3}rem; padding-right: {4}rem;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine()

        $canvasHeight = Get-HtmlCanvasHeight -Orientation $Orientation
        [ref] $null = $pageBreakBuilder.AppendFormat('<div style="min-height: {0}mm" >', $canvasHeight)
        [ref] $null = $pageBreakBuilder.Append((Out-HtmlHeaderFooter -Header))

        return $pageBreakBuilder.ToString()
    }
}

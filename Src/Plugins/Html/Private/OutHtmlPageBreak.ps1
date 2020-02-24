function OutHtmlPageBreak
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
        [System.Text.StringBuilder] $pageBreakBuilder = New-Object 'System.Text.StringBuilder';
        [ref] $null = $pageBreakBuilder.Append('</div></div>');
        $topMargin = ConvertMmToEm $Document.Options['MarginTop'];
        $leftMargin = ConvertMmToEm $Document.Options['MarginLeft'];
        $bottomMargin = ConvertMmToEm $Document.Options['MarginBottom'];
        $rightMargin = ConvertMmToEm $Document.Options['MarginRight'];
        [ref] $null = $pageBreakBuilder.AppendFormat('<div class="{0}">', $Orientation.ToLower());
        [ref] $null = $pageBreakBuilder.AppendFormat('<div class="{0}" style="padding-top: {1}rem; padding-left: {2}rem; padding-bottom: {3}rem; padding-right: {4}rem;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
        return $pageBreakBuilder.ToString();
    }
}

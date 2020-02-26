function OutTextPageBreak
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
        $lineBreak = OutTextLineBreak
        $pageBreak = '{0}{1}' -f $lineBreak, [System.Environment]::NewLine
        return $pageBreak
    }
}

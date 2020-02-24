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
        return "$(OutTextLineBreak)`r`n";
    }
}

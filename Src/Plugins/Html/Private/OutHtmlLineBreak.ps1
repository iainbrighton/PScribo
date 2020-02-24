function OutHtmlLineBreak
{
<#
    .SYNOPSIS
        Output formatted Html line break.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param ( )
    process
    {
        return '<hr />';
    }
}

function OutTextLineBreak
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
            $options = New-PScriboTextOption
        }
    }
    process
    {
        ## Use the specified output width
        if ($options.TextWidth -eq 0)
        {
            $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1
        }
        $lb = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator)
        return "$(OutStringWrap -InputObject $lb -Width $options.TextWidth)`r`n"
    }
}

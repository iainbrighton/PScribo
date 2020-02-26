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
        $padding = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator)
        $lineBreak = '{0}{1}' -f (OutStringWrap -InputObject $padding -Width $options.TextWidth), [System.Environment]::NewLine
        return $lineBreak
    }
}

function Out-TextLineBreak
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
        $convertToAlignedStringParams = @{
            InputObject = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator)
            Width       = $options.TextWidth
        }
        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

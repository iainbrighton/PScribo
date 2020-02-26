function OutStringWrap
{
<#
    .SYNOPSIS
        Outputs objects to strings, wrapping as required.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [Object[]] $InputObject,

        [Parameter()]
        [ValidateNotNull()]
        [System.Int32] $Width = $Host.UI.RawUI.BufferSize.Width
    )
    begin
    {
        ## 2 is the minimum, therefore default to wiiiiiiiiiide!
        if ($Width -lt 2)
        {
            $Width = 4096
        }
        WriteLog -Message ('Wrapping text at "{0}" characters.' -f $Width) -IsDebug
    }
    process
    {
        foreach ($object in $InputObject)
        {
            $textBuilder = New-Object -TypeName System.Text.StringBuilder
            $text = (Out-String -InputObject $object).TrimEnd([System.Environment]::NewLine)
            for ($i = 0; $i -le $text.Length; $i += $Width)
            {
                if (($i + $Width) -ge ($text.Length -1))
                {
                    [ref] $null = $textBuilder.Append($text.Substring($i))
                }
                else
                {
                    [ref] $null = $textBuilder.AppendLine($text.Substring($i, $Width))
                }
            }
            return $textBuilder.ToString()
            $textBuilder = $null
        }
    }
}

function Out-TextImage
{
<#
    .SYNOPSIS
        Output formatted image text.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Image
    )
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
            InputObject = '[Image Text="{0}"]' -f $Image.Text
            Width       = $options.TextWidth
        }

        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

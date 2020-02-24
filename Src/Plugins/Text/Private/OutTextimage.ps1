function OutTextimage
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
            $options = New-PScriboTextOption;
        }
    }
    process
    {
        $imageText = '[Image Text="{0}"]' -f $Image.Text
        return OutStringWrap -InputObject $imageText -Width $Options.TextWidth
    }
}

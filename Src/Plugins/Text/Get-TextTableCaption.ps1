function Get-TextTableCaption
{
<#
    .SYNOPSIS
        Generates html <p> caption from a PScribo.Table object.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
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
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style
        $convertToAlignedStringParams = @{
            InputObject = '{0} {1} {2}' -f $tableStyle.CaptionPrefix, $Table.Number, $Table.Caption
            Width       = $options.TextWidth
            Tabs        = $Table.Tabs
            Align       = $tableStyle.Align
        }
        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

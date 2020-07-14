function BlankLine {
<#
    .SYNOPSIS
        Initializes a new PScribo blank line object.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline, Position = 0)]
        [System.UInt32] $Count = 1
    )
    process
    {
        Write-PScriboMessage -Message $localized.ProcessingBlankLine
        return (New-PScriboBlankLine @PSBoundParameters)
    }
}

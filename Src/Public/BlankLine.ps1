function BlankLine {
<#
    .SYNOPSIS
        Initializes a new PScribo blank line object.
#>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [System.UInt32] $Count = 1
    )
    begin {
        <#! BlankLine.Internal.ps1 !#>
    } #end begin
    process {

        WriteLog -Message $localized.ProcessingBlankLine;
        return (New-PScriboBlankLine @PSBoundParameters);

    } #end process
} #end function BlankLine

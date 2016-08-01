function LineBreak {
<#
    .SYNOPSIS
        Initializes a new PScribo line break object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {

        <#! LineBreak.Internal.ps1 !#>

    } #end begin
    process {

        WriteLog -Message $localized.ProcessingLineBreak;
        return (New-PScriboLineBreak @PSBoundParameters);

    } #end process
} #end function LineBreak

function PageBreak {
<#
    .SYNOPSIS
        Creates a PScribo page break object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {

        <#! PageBreak.Internal.ps1 !#>

    } #end begin
    process {

        WriteLog -Message $localized.ProcessingPageBreak;
        return (New-PScriboPageBreak -Id $Id);

    }
} #end function PageBreak

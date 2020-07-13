function PageBreak {
<#
    .SYNOPSIS
        Creates a PScribo page break object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    process
    {
        Write-PScriboMessage -Message $localized.ProcessingPageBreak
        return (New-PScriboPageBreak -Id $Id)
    }
}

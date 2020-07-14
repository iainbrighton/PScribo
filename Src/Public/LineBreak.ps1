function LineBreak {
<#
    .SYNOPSIS
        Initializes a new PScribo line break object.
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
        Write-PScriboMessage -Message $localized.ProcessingLineBreak
        return (New-PScriboLineBreak @PSBoundParameters)
    }
}

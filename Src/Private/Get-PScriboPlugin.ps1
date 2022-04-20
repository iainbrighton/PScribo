function Get-PScriboPlugin
{
<#
    .SYNOPSIS
        Returns available PScribo plugins.
#>
    [CmdletBinding()]
    param ( )
    process
    {
        Get-ChildItem -Path (Join-Path -Path $pscriboRoot -ChildPath '\Src\Plugins') |
            Select-Object -ExpandProperty Name
    }
}

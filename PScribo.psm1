Set-StrictMode -Version Latest

## Import localisation strings based on UICulture
$importLocalizedDataParams = @{
    BindingVariable = 'localized'
    BaseDirectory   = $PSScriptRoot
    FileName        = 'PScribo.Resources.psd1'
}
Import-LocalizedData @importLocalizedDataParams -ErrorAction SilentlyContinue

#Fallback to en-US culture strings
if (-not (Test-Path -Path 'Variable:\localized'))
{
    $importLocalizedDataParams['UICulture'] = 'en-US'
    Import-LocalizedData @importLocalizedDataParams -ErrorAction Stop
}

## Dot source all the nested .ps1 files in the \Functions and \Plugin folders, excluding tests
$pscriboRoot = Split-Path -Parent $PSCommandPath
Get-ChildItem -Path "$pscriboRoot\Src\" -Include '*.ps1' -Recurse |
    ForEach-Object {
        Write-Debug ($localized.ImportingFile -f $_.FullName)
        ## https://becomelotr.wordpress.com/2017/02/13/expensive-dot-sourcing/
        . ([System.Management.Automation.ScriptBlock]::Create(
                [System.IO.File]::ReadAllText($_.FullName)
            ))
    }

Add-Type -AssemblyName 'System.Drawing'
#Export-ModuleMember -Function $exportedFunctions -Alias $exportedAliases -Verbose:$false

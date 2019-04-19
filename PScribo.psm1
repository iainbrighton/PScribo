Set-StrictMode -Version Latest;

## Import localisation strings based on UICulture
Import-LocalizedData -BindingVariable localized -BaseDirectory $PSScriptRoot -FileName PScribo.Resources.psd1 -ErrorAction SilentlyContinue

#Fallback to en-US culture strings
if (-not (Test-Path -Path 'Variable:\localized')) {
    Import-LocalizedData -BaseDirectory $PSScriptRoot -BindingVariable localized -UICulture 'en-US' -FileName PScribo.Resources.psd1 -ErrorAction Stop
}

## Dot source all the nested .ps1 files in the \Functions and \Plugin folders, excluding tests
$pscriboRoot = Split-Path -Parent $PSCommandPath;
Get-ChildItem -Path "$pscriboRoot\Src\" -Include '*.ps1' -Recurse |
    ForEach-Object {
        Write-Verbose ($localized.ImportingFile -f $_.FullName);
        ## https://becomelotr.wordpress.com/2017/02/13/expensive-dot-sourcing/
        . ([System.Management.Automation.ScriptBlock]::Create(
                [System.IO.File]::ReadAllText($_.FullName)
            ));
    }

$exportedFunctions = @(
    'Document',
    'Export-Document',
    'Section',
    'DocumentOption',
    'LineBreak',
    'PageBreak',
    'Paragraph',
    'Image',
    'Section',
    'Style',
    'Table',
    'TableStyle',
    'Set-Style',
    'TOC',
    'BlankLine'
);

$exportedAliases = @(
    'GlobalOption'
);

Add-Type -AssemblyName 'System.Drawing'

Export-ModuleMember -Function $exportedFunctions -Alias $exportedAliases;

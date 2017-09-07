Set-StrictMode -Version Latest;

## Import localisation strings
$importLocalizedDataParams = @{
    BindingVariable = 'localized';
    FileName = 'PScribo.Resources.psd1';
    BaseDirectory = $PSScriptRoot;
}
Import-LocalizedData @importLocalizedDataParams;

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

Export-ModuleMember -Function $exportedFunctions -Alias $exportedAliases;

## Load the System.Drawing.dll ## TODO: THIS WON'T WORK ON .NET CORE :|
$systemDrawingPath = Get-ChildItem -Path "$env:SystemRoot\assembly" -Filter *drawing.dll -Recurse | Select-Object -First 1
$null = [Reflection.Assembly]::LoadFrom($systemDrawingPath.FullName)

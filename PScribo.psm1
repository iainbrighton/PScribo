Set-StrictMode -Version Latest;

## Import localisation strings
$importLocalizedDataParams = @{
    BindingVariable = 'localized';
    FileName = 'PScribo.Resources.psd1';
    BaseDirectory = $PSScriptRoot;
}
if (-not (Test-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath $PSUICulture))) {

    # fallback to en-US
    $importLocalizedDataParams['UICulture'] = 'en-US';
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

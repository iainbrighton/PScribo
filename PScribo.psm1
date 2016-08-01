Set-StrictMode -Version Latest;
Import-LocalizedData -BindingVariable localized -FileName Resources.psd1

## Dot source all the nested .ps1 files in the \Functions and \Plugin folders, excluding tests
$pscriboRoot = Split-Path -Parent $PSCommandPath;
Get-ChildItem -Path "$pscriboRoot\Src\" -Include '*.ps1' -Recurse |
    ForEach-Object {
        Write-Verbose ($localized.ImportingFile -f $_.FullName);
        . $_.FullName;
    }

$exportedFunctions = @(
  'Document',
  'Export-Document',
  'Section',
  'GlobalOption',
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

Export-ModuleMember -Function $exportedFunctions;

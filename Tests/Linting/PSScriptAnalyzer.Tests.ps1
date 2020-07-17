#requires -Version 4
#requires -Modules PSScriptAnalyzer

$repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path;
Describe 'Linting\PSScriptAnalyzer' {

    Get-ChildItem -Path "$repoRoot\Src\*.ps*1" -Recurse -File | ForEach-Object {
        It "File '$($_.FullName.Replace($repoRoot,''))' passes PSScriptAnalyzer rules" {
            $result = Invoke-ScriptAnalyzer -Path $_.FullName -Severity Warning | Select-Object -ExpandProperty Message
            $result | Should BeNullOrEmpty
        }
    }
}

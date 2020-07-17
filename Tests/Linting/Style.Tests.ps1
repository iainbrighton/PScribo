$repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path;

#Style rules based on Pester v. 4.0.2-rc2
Describe 'Linting\Style' -Tags "Style" {

    $excludedPaths = @(
                        '.git*',
                        '.vscode',
                        'DSCResources', # We'll take the public DSC resources as-is
                        'Release',
                        '*.png',
                        '*.jpg',
                        '*.docx',
                        '*.enc',
                        '*.dll',
                        'appveyor-tools',
                        'TestResults.xml',
                        'Tests',
                        'Docs',
                        'PScriboExample.txt',
                        'Lib'
                    );

    function TestStylePath {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [System.String] $Path,

            [System.String[]] $Exclude
        )
        process
        {
            Get-ChildItem -Path $Path -Exclude $Exclude |
                ForEach-Object {
                    if ($_ -is [System.IO.FileInfo])
                    {
                        It "File '$($_.FullName.Replace($repoRoot,''))' contains no trailing whitespace" {
                            $badLines = @(
                                $lines = [System.IO.File]::ReadAllLines($_.FullName)
                                $lineCount = $lines.Count

                                for ($i = 0; $i -lt $lineCount; $i++) {
                                    if ($lines[$i] -match '\s+$') {
                                        'File: {0}, Line: {1}' -f $_.FullName, ($i + 1)
                                    }
                                }
                            )

                            @($badLines).Count | Should Be 0
                        }

                        It "File '$($_.FullName.Replace($repoRoot,''))' ends with a newline" {

                            $string = [System.IO.File]::ReadAllText($_.FullName)
                            ($string.Length -gt 0 -and $string[-1] -ne "`n") | Should Be $false
                        }
                    }
                    elseif ($_ -is [System.IO.DirectoryInfo])
                    {
                        TestStylePath -Path $_.FullName -Exclude $Exclude
                    }
                }
        } #end process
    } #end function

    Get-ChildItem -Path $repoRoot -Exclude $excludedPaths |
        ForEach-Object {
            TestStylePath -Path $_.FullName -Exclude $excludedPaths
        }
}

<#
$ModuleName = "PScribo"

#Provided path assume that your module manifest (a file with the psd1 extension) exists in the parent directory for directory where test script is stored
$RelativePathToModuleRoot = "{0}\..\..\..\{1}\" -f $PSScriptRoot, $ModuleName
$RelativePathToModuleManifest = "{0}\{1}.psd1" -f $RelativePathToModuleRoot, $ModuleName


#Style rules based on Pester v. 4.0.2-rc2
Describe 'Style rules' -Tags "Style"{

    $files = @(
        Get-ChildItem (Join-Path $RelativePathToModuleRoot '.\Src') -Include *.ps1, *.psm1 -Recurse
        Get-ChildItem $RelativePathToModuleRoot\* -Include *.ps1, *.psm1
        Get-ChildItem (Join-Path $RelativePathToModuleRoot '.\Tests') -Include *.ps1, *.psm1, *.psd1 -Recurse
        Get-ChildItem (Join-Path $RelativePathToModuleRoot '.\en-US') -Include *.help.txt -Recurse
    )

    It "$ModuleName source files contain no trailing whitespace" {
        $badLines = @(
            foreach ($file in $files) {
                $lines = [System.IO.File]::ReadAllLines($file.FullName)
                $lineCount = $lines.Count

                for ($i = 0; $i -lt $lineCount; $i++) {
                    if ($lines[$i] -match '\s+$') {
                        'File: {0}, Line: {1}' -f $file.FullName, ($i + 1)
                    }
                }
            }
        )

        if ($badLines.Count -gt 0) {
            throw "The following $($badLines.Count) lines contain trailing whitespace: `r`n`r`n$($badLines -join "`r`n")"
        }
    }

    It "$ModuleName source files lines start with a tab character" {
        $badLines = @(
            foreach ($file in $files) {
                $lines = [System.IO.File]::ReadAllLines($file.FullName)
                $lineCount = $lines.Count

                for ($i = 0; $i -lt $lineCount; $i++) {
                    if ($lines[$i] -match '^[  ]*\t|^\t|^\t[  ]*') {
                        'File: {0}, Line: {1}' -f $file.FullName, ($i + 1)
                    }
                }
            }
        )

        if ($badLines.Count -gt 0) {
            throw "The following $($badLines.Count) lines start with a tab character: `r`n`r`n$($badLines -join "`r`n")"
        }
    }

    It "$ModuleName source files all end with a newline" {
        $badFiles = @(
            foreach ($file in $files) {
                $string = [System.IO.File]::ReadAllText($file.FullName)
                if ($string.Length -gt 0 -and $string[-1] -ne "`n") {
                    $file.FullName
                }
            }
        )

        if ($badFiles.Count -gt 0) {
            throw "The following files do not end with a newline: `r`n`r`n$($badFiles -join "`r`n")"
        }
    }
}
#>

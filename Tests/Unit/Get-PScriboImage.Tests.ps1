$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    $testRoot = Split-Path -Path $PSScriptRoot -Parent

    Describe -Name 'Get-PScriboImage' -Fixture {

        $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock {
            Image -Path (Join-Path -Path $testRoot -ChildPath 'TestImage.jpg') -Id 1
            Section Nested {
                Image -Path (Join-Path -Path $testRoot -ChildPath 'TestImage.png') -Id 2
            }
        }

        It 'finds all Images' {
            $result = Get-PScriboImage -Section $pscriboDocument.Sections

            $result.Count | Should Be 2
        }

        It 'finds Image by Id' {
            $result = @(Get-PScriboImage -Section $pscriboDocument.Sections -Id 2)

            $result.Count | Should Be 1
        }
    }
}

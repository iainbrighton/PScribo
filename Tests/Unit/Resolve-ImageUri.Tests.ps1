$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    Describe -Name 'Resolve-ImageUri' -Fixture {

        It "returns 'System.Uri' type" {
            $testPath = 'about:Blank'

            $result = Resolve-ImageUri -Path $testPath

            $result -is [System.Uri] | Should Be $true
        }
    }
}

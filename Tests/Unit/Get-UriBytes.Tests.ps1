$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    $testRoot = Split-Path -Path $PSScriptRoot -Parent;

    Describe -Name 'Get-UriBytes' -Fixture {

        It "returns 'System.Byte[]' type" {

            $testPath = Join-Path -Path $testRoot -ChildPath 'TestImage.jpg'
            $testUri = Resolve-ImageUri -Path $testPath

            $result = Get-UriBytes -Uri $testUri

            $result -is [System.Object[]] | Should Be $true
            $result[0] -is [System.Byte] | Should Be $true
        }

    }
}

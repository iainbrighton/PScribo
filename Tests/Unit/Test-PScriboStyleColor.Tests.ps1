$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

Describe 'Test-PScriboStyleColor' {

        It 'tests valid html color code' {
            Test-PscriboStyleColor -Color '012345' | Should Be $true
            Test-PscriboStyleColor -Color '#012345' | Should Be $true
            Test-PscriboStyleColor -Color '#D3D' | Should Be $true
            Test-PscriboStyleColor -Color D3D | Should Be $true
        }

        It 'tests invalid length html color code' {
            Test-PscriboStyleColor -Color abcd | Should Be $false
            Test-PscriboStyleColor -Color 1abcdef | Should Be $false
        }

        It 'tests invalid html color code' {
            Test-PscriboStyleColor -Color ghi | Should Be $false
        }
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    $testRoot = Split-Path -Path $PSScriptRoot -Parent

    Describe -Name 'ConvertTo-Image' -Fixture {

        It "returns 'System.Drawing.Image' type" {
            $testPath = Join-Path -Path $testRoot -ChildPath 'TestImage.jpg'
            $testUri = Resolve-ImageUri -Path $testPath
            $testBytes = Get-UriBytes -Uri $testUri

            $result = ConvertTo-Image -Bytes $testBytes

            $result -is [System.Drawing.Image] | Should Be $true
        }

    }
}

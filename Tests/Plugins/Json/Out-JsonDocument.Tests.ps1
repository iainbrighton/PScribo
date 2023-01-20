$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Json\Out-JsonDocument' {

        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'calls Out-JsonSection' {
            Mock -CommandName Out-JsonSection

            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | Out-JsonDocument -Path $path

            Assert-MockCalled -CommandName Out-JsonSection -Exactly 1
        }

        It 'calls Out-JsonParagraph' {
            Mock -CommandName Out-JsonParagraph

            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | Out-JsonDocument -Path $path

            Assert-MockCalled -CommandName Out-JsonParagraph -Exactly 1
        }

         It 'calls Out-JsonTable' {
            Mock -CommandName Out-JsonTable

            Document -Name 'TestDocument' -ScriptBlock { Get-Process | Select-Object -First 1 | Table 'TestTable' } | Out-JsonDocument -Path $path

            Assert-MockCalled -CommandName Out-JsonTable -Exactly 1
        }

        It 'calls Out-JsonTOC' {
            Mock -CommandName Out-JsonTOC

            Document -Name 'TestDocument' -ScriptBlock { TOC -Name 'TestTOC'; } | Out-JsonDocument -Path $path

            Assert-MockCalled -CommandName Out-JsonTOC -Exactly 1
        }
    }
}

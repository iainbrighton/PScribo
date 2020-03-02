$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    $isNix = $false
    if (($PSVersionTable['PSEdition'] -eq 'Core') -and (-not $IsWindows))
    {
        $isNix = $true
    }

    Describe 'Plugins\Text\Out-TextDocument' {

        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'calls Out-TextSection' {
            Mock -CommandName Out-TextSection

            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextSection -Exactly 1
        }

        It 'calls Out-TextParagraph' {
            Mock -CommandName Out-TextParagraph

            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextParagraph -Exactly 1
        }

        It 'calls Out-TextLineBreak' {
            Mock -CommandName Out-TextLineBreak

            Document -Name 'TestDocument' -ScriptBlock { LineBreak; } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextLineBreak -Exactly 1
        }

        It 'calls Out-TextPageBreak' {
            Mock -CommandName Out-TextPageBreak

            Document -Name 'TestDocument' -ScriptBlock { PageBreak; } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextPageBreak -Exactly 1
        }

         It 'calls Out-TextTable' {
            Mock -CommandName Out-TextTable

            Document -Name 'TestDocument' -ScriptBlock { Get-Process | Select-Object -First 1 | Table 'TestTable' } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextTable -Exactly 1
        }

        It 'calls Out-TextTOC' {
            Mock -CommandName Out-TextTOC

            Document -Name 'TestDocument' -ScriptBlock { TOC -Name 'TestTOC'; } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextTOC -Exactly 1
        }

        It 'calls Out-TextBlankLine' {
            Mock -CommandName Out-TextBlankLine

            Document -Name 'TestDocument' -ScriptBlock { BlankLine; } | Out-TextDocument -Path $path

            Assert-MockCalled -CommandName Out-TextBlankLine -Exactly 1
        }

        It 'calls Out-TextBlankLine twice' {
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; BlankLine; } | Out-TextDocument -Path $path;

            Assert-MockCalled -CommandName Out-TextBlankLine -Exactly 3 ## Mock calls are cumalative
        }
    }
}

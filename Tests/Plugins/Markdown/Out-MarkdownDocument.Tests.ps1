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

    Describe 'Plugins\Markdown\Out-MarkdownDocument' {

        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'calls Out-MarkdownSection' {
            Mock -CommandName Out-MarkdownSection

            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownSection -Exactly 1
        }

        It 'calls Out-MarkdownParagraph' {
            Mock -CommandName Out-MarkdownParagraph

            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownParagraph -Exactly 1
        }

        It 'calls Out-MarkdownLineBreak' {
            Mock -CommandName Out-MarkdownLineBreak

            Document -Name 'TestDocument' -ScriptBlock { LineBreak } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownLineBreak -Exactly 1
        }

        It 'calls Out-MarkdownPageBreak' {
            Mock -CommandName Out-MarkdownPageBreak

            Document -Name 'TestDocument' -ScriptBlock { PageBreak } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownPageBreak -Exactly 1
        }

         It 'calls Out-MarkdownTable' {
            Mock -CommandName Out-MarkdownTable

            Document -Name 'TestDocument' -ScriptBlock { Get-Process | Select-Object -First 1 | Table 'TestTable' } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownTable -Exactly 1
        }

        It 'calls Out-MarkdownTOC' {
            Mock -CommandName Out-MarkdownTOC

            Document -Name 'TestDocument' -ScriptBlock { TOC -Name 'TestTOC' } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownTOC -Exactly 1
        }

        It 'calls Out-MarkdownBlankLine' {
            Mock -CommandName Out-MarkdownBlankLine

            Document -Name 'TestDocument' -ScriptBlock { BlankLine } | Out-MarkdownDocument -Path $path

            Assert-MockCalled -CommandName Out-MarkdownBlankLine -Exactly 1
        }

        It 'calls Out-MarkdownBlankLine twice' {
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; BlankLine } | Out-MarkdownDocument -Path $path;

            Assert-MockCalled -CommandName Out-MarkdownBlankLine -Exactly 3 ## Mock calls are cumalative
        }
    }
}

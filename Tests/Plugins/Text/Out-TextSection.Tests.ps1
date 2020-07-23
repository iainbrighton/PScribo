$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextSection' {

        $Document = Document -Name 'TestDocument' -ScriptBlock { }
        $pscriboDocument = $Document

        It 'outputs indented section (#73)' {
            $result = Section -Name TestSection -ScriptBlock { } -Tabs 2 | Out-TextSection

            $result -match '^\r?\n        ' | Should Be $true
        }

        It 'calls Out-TextParagraph' {
            Mock -CommandName Out-TextParagraph

            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextParagraph -Exactly 1
        }

        It 'calls Out-TextParagraph twice' {
            Mock -CommandName Out-TextParagraph

            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph'; Paragraph 'TestParagraph'; } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextParagraph -Exactly 3
        }

        It 'calls Out-TextTable' {
            Mock -CommandName Out-TextTable

            Section -Name TestSection -ScriptBlock { Get-Process | Select-Object -First 3 | Table TestTable } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextTable -Exactly 1
        }

        It 'calls Out-TextPageBreak' {
            Mock -CommandName Out-TextPageBreak

            Section -Name TestSection -ScriptBlock { PageBreak } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextPageBreak -Exactly 1
        }

         It 'calls Out-TextLineBreak' {
            Mock -CommandName Out-TextLineBreak

            Section -Name TestSection -ScriptBlock { LineBreak } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextLineBreak -Exactly 1
        }

        It 'calls Out-TextBlankLine' {
            Mock -CommandName Out-TextBlankLine

            Section -Name TestSection -ScriptBlock { BlankLine } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextBlankLine -Exactly 1
        }

        It 'warns on call Out-TextTOC' {
            Mock -CommandName Out-TextTOC

            Section -Name TestSection -ScriptBlock { TOC 'TestTOC' } | Out-TextSection -WarningAction SilentlyContinue

            Assert-MockCalled Out-TextTOC -Exactly 0
        }

        It 'calls nested Out-TextSection' {
            ## Note this must be called last in the Describe script block as the Out-XmlSection gets mocked!
            Mock -CommandName Out-TextSection

            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | Out-TextSection

            Assert-MockCalled -CommandName Out-TextSection -Exactly 1
        }
    }
}

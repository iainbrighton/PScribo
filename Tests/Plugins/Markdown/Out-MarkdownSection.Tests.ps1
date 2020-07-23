$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Markdown\Out-MarkdownSection' {

        $Document = Document -Name 'TestDocument' -ScriptBlock { }
        $pscriboDocument = $Document
        $script:currentPScriboObject = 'PScribo.Document'

        It 'calls Out-MarkdownParagraph' {
            Mock -CommandName Out-MarkdownParagraph

            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownParagraph -Exactly 1;
        }

        It 'calls Out-MarkdownParagraph twice' {
            Mock -CommandName Out-MarkdownParagraph

            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph'; Paragraph 'TestParagraph'; } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownParagraph -Exactly 3
        }

        It 'calls Out-MarkdownTable' {
            Mock -CommandName Out-MarkdownTable

            Section -Name TestSection -ScriptBlock { Get-Process | Select-Object -First 3 | Table TestTable } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownTable -Exactly 1
        }

        It 'calls Out-MarkdownPageBreak' {
            Mock -CommandName Out-MarkdownPageBreak

            Section -Name TestSection -ScriptBlock { PageBreak } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownPageBreak -Exactly 1
        }

         It 'calls Out-MarkdownLineBreak' {
            Mock -CommandName Out-MarkdownLineBreak

            Section -Name TestSection -ScriptBlock { LineBreak } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownLineBreak -Exactly 1
        }

        It 'calls Out-MarkdownBlankLine' {
            Mock -CommandName Out-MarkdownBlankLine

            Section -Name TestSection -ScriptBlock { BlankLine } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownBlankLine -Exactly 1
        }

        It 'warns on call Out-MarkdownTOC' {
            Mock -CommandName Out-MarkdownTOC

            Section -Name TestSection -ScriptBlock { TOC 'TestTOC' } | Out-MarkdownSection -WarningAction SilentlyContinue

            Assert-MockCalled Out-MarkdownTOC -Exactly 0
        }

        It 'calls nested Out-MarkdownSection' {
            ## Note this must be called last in the Describe script block as the Out-XmlSection gets mocked!
            Mock -CommandName Out-MarkdownSection

            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | Out-MarkdownSection

            Assert-MockCalled -CommandName Out-MarkdownSection -Exactly 1
        }
    }
}

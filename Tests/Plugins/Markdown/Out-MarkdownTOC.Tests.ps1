$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Markdown\Out-MarkdownTOC' {

        It 'outputs TOC name' {
            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading2 = 'Heading 2'
            $Document = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 { }
                Section $heading2 -Style Heading1 { }
            }
            $expected = '# {0}{1}{1}' -f $tocName, [System.Environment]::NewLine

            $result = Out-MarkdownTOC -TOC $Document.Sections[0]

            $result | Should MatchExactly $expected
        }

        It 'outputs TOC section heading and anchor' {
            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading1anchor = 'heading-1'
            $heading2 = 'Heading 2'
            $heading2anchor = 'heading-2'
            $Document = Document -Name 'TestDocument' -ScriptBlock {
                TOC -Name $tocName
                Section $heading1 -Style Heading1 { }
                Section $heading2 -Style Heading1 { }
            }
            $result = Out-MarkdownTOC -TOC $Document.Sections[0]

            $result | Should Match ('\[{0}\]\(#{1}\)' -f $heading1, $heading1anchor)
            $result | Should Match ('\[{0}\]\(#{1}\)' -f $heading2, $heading2anchor)
        }

        It 'outputs TOC section heading number and anchor' {
            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading1anchor = 'heading-1'
            $heading11 = 'Heading 1.1'
            $heading11anchor = 'heading-1.1'
            $heading2 = 'Heading 2'
            $heading2anchor = 'heading-2'

            $Document = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading11 -Style Heading1 { }
                }
                Section $heading2 -Style Heading1 { }
            }
            $options = Merge-PScriboPluginOption -DocumentOptions $Document.Options -PluginOptions (New-PScriboTextOption)
            $result = Out-MarkdownTOC -TOC $Document.Sections[0]

            $result | Should Match ('\[1 {0}\]\(#1-{1}\)' -f $heading1, $heading1anchor)
            $result | Should Match ('\[1.1 {0}\]\(#1.1-{1}\)' -f $heading11, $heading11anchor)
            $result | Should Match ('\[2 {0}\]\(#2-{1}\)' -f $heading2, $heading2anchor)
        }

    }
}

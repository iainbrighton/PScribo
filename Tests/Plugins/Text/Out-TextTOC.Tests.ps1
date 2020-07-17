$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextTOC' {

        It 'outputs TOC name and section break' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading2 = 'Heading 2'
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{1}-+' -f $tocName, [System.Environment]::NewLine

            $result = Out-TextTOC -TOC $Document.Sections[0]

            $result | Should MatchExactly $expected
        }

        It 'adds section numbers when "EnableSectionNumbering" is enabled (#20)' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading2 = 'Heading 2'
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{3}-+{3}1\s+{1}{3}1.1\s+{2}{3}$' -f $tocName, $heading1, $heading2, [System.Environment]::NewLine

            $options = Merge-PScriboPluginOption -DocumentOptions $Document.Options -PluginOptions (New-PScriboTextOption)
            $result = Out-TextTOC -TOC $Document.Sections[0] -Verbose

            $result | Should Match $expected
        }

        It 'does not add section numbers when "EnableSectionNumbering" is disabled (#20)' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading2 = 'Heading 2'
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{3}-+{3}{1}{3} {2}{3}$' -f $tocName, $heading1, $heading2, [System.Environment]::NewLine

            $result = Out-TextTOC -TOC $Document.Sections[0] -Verbose

            $result | Should Match $expected
        }
    }
}

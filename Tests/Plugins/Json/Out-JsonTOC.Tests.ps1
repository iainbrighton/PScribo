$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Json\Out-JsonTOC' {

        ## Section numbering is ignored and used regardless
        
        It 'outputs TOC' {
            $tocName = 'Table of contents'
            $heading1 = 'Heading 1'
            $heading2 = 'Heading 2'
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '{{.*"{0}".*"{1}"}}' -f $heading1, $heading2

            $result = Out-JsonTOC -TOC $Document.Sections[0] | ConvertTo-Json -Depth 100 -Compress

            $result | Should MatchExactly $expected
        }
    }
}

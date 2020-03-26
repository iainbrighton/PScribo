$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlPageBreak' {

        ## Scaffold new document to initialise options/styles
        $Document = Document -Name 'Test' -ScriptBlock { }
        $script:currentPageNumber = 1
        $text = Out-HtmlPageBreak -Orientation Portrait

        It 'closes previous </div> tags' {
            $text.StartsWith('</div></div>') | Should Be $true
        }

        It 'creates new <div class="portrait">' {
            $text -match '<div class="portrait">' | Should Be $true
        }

        It 'sets page class to default style' {
            $divStyleMatch = '<div class="{0}"' -f $Document.DefaultStyle
            $text -match $divStyleMatch | Should Be $true
        }

        It 'includes page margins' {
            $text -match 'padding-top:[\s\S]+em' | Should Be $true
            $text -match 'padding-right:[\s\S]+em' | Should Be $true
            $text -match 'padding-bottom:[\s\S]+em' | Should Be $true
            $text -match 'padding-left:[\s\S]+em' | Should Be $true
        }
    }
}

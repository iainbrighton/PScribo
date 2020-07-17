$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlBlankLine' {

        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        It 'creates a single <br /> html tag' {
            BlankLine | Out-HtmlBlankLine | Should BeExactly '<br />'
        }

        It 'creates two <br /> html tags' {
            BlankLine -Count 2 | Out-HtmlBlankLine | Should BeExactly '<br /><br />'
        }
    }
}
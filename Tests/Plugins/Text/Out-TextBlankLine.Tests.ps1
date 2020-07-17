$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextBlankLine' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        It 'Defaults to a single blank line' {
            $expected = [System.Environment]::NewLine

            $l = BlankLine | Out-TextBlankLine

            $l | Should Be $expected
        }

        It 'Creates 3 blank lines' {
            $expected = '{0}{0}{0}' -f [System.Environment]::NewLine

            $l = BlankLine -Count 3 | Out-TextBlankLine

            $l | Should Be $expected
        }
    }
}

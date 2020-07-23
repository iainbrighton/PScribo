$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Markdown\Out-MarkdownBlankLine' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        It 'defaults to a single blank line' {
            $expected = [System.Environment]::NewLine

            $l = BlankLine |Out-MarkdownBlankLine

            $l | Should Be $expected
        }

        It 'creates 3 blank lines' {
            $expected = '{0}{0}{0}' -f [System.Environment]::NewLine

            $l = BlankLine -Count 3 | Out-MarkdownBlankLine

            $l | Should Be $expected
        }

        It 'outputs Html line breaks when "RenderBlankLine" option is enabled' {
            $options = New-PScriboMarkdownOption -RenderBlankLine $true
            $expected = '<br />{0}<br />{0}<br />{0}' -f [System.Environment]::NewLine

            $l = BlankLine -Count 3 | Out-MarkdownBlankLine

            $l | Should Be $expected
        }

    }
}

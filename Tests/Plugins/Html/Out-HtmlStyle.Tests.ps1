$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlStyle' {

        It 'creates <style> tag' {
            $Document = Document -Name 'Test' -ScriptBlock { }
            $text = Out-HtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles

            $text -match '<style type="text/css">' | Should Be $true
            $text -match '</style>' | Should Be $true
        }

        It 'creates page layout style by default' {
            $Document = Document -Name 'Test' -ScriptBlock { }
            $text = Out-HtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles

            $text -match 'html {' | Should Be $true
            $text -match 'page {' | Should Be $true
            $text -match '@media print {' | Should Be $true
        }

        It "suppresses page layout style when 'Options.NoPageLayoutSyle' specified" {
            $Document = Document -Name 'Test' -ScriptBlock { }
            $text = Out-HtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles -NoPageLayoutStyle

            $text -match 'html {' | Should Be $false
            $text -match 'page {' | Should Be $false
            $text -match '@media print {' | Should Be $false
        }
    }
}

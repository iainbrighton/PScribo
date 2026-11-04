$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginsRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginsRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Word\Get-WordNumberingDocument' {

        It 'adds no section heading list by default' {
            $document = Document -Name 'TestDocument' {
                List -Item 'Item' -Numbered
            }

            $numberingXml = Get-WordNumberingDocument -Lists $document.Lists

            $numberingXml.DocumentElement.SelectNodes('*[local-name()="num"]').Count | Should Be 1
        }

        It 'adds the section heading list after the document lists when enabled' {
            $document = Document -Name 'TestDocument' {
                List -Item 'Item' -Numbered
            }

            $numberingXml = Get-WordNumberingDocument -Lists $document.Lists -EnableWordSectionNumbering

            $numberingXml.DocumentElement.OuterXml | Should Match '<w:num w:numId="2"><w:abstractNumId w:val="1" /></w:num>'
        }

        It 'numbers the section heading levels "%1.%2.%3"' {
            $document = Document -Name 'TestDocument' { }

            $numberingXml = Get-WordNumberingDocument -Lists $document.Lists -EnableWordSectionNumbering

            $numberingXml.DocumentElement.OuterXml | Should Match '<w:lvl w:ilvl="2">[\s\S]+?<w:lvlText w:val="%1.%2.%3" />'
        }
    }
}

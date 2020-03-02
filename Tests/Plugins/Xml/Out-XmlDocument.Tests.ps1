$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    ## Scaffold root XmlDocument
    $xmlDocument = New-Object -TypeName System.Xml.XmlDocument;
    [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
    $element = $xmlDocument.AppendChild($xmlDocument.CreateElement('testdocument'));
    [ref] $null = $element.SetAttribute('name', 'testdocument');

    Describe 'Plugins\Text\Out-XmlDocument' {

        It 'calls Out-XmlTable' {
            Mock -CommandName Out-XmlTable -MockWith { return $xmlDocument.CreateElement('TestTable') }

            Document -Name 'TestDocument' -ScriptBlock { Get-Process | Select-Object -First 1 | Table 'TestTable' } | Out-XmlDocument -Path $testDrive

            Assert-MockCalled -CommandName Out-XmlTable -Exactly 1
        }

        It 'calls Out-XmlParagraph' {
            Mock -CommandName Out-XmlParagraph -MockWith { return $xmlDocument.CreateElement('TestParagraph') }

            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | Out-XmlDocument -Path $testDrive

            Assert-MockCalled -CommandName Out-XmlParagraph -Exactly 1
        }

        It 'calls Out-XmlSection' {
            Mock -CommandName Out-XmlSection -MockWith { return $xmlDocument.CreateElement('TestSection'); }

            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | Out-XmlDocument -Path $testDrive

            Assert-MockCalled -CommandName Out-XmlSection -Exactly 1
        }
    }
}

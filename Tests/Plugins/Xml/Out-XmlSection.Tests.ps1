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

    ## Scaffold document
    $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {
        Paragraph -Name 'TestParagraph';
        Get-Process | Select-Object -First 3 | Table 'TestTable';
        Section -Name 'TestSection' -ScriptBlock {
            Section -Name 'TestSubSection' { }
        }
    }

    Describe 'Plugins\Xml\Out-XmlSection' {

        It 'calls Out-XmlTable' {
            Mock -CommandName Out-XmlTable -MockWith { return $xmlDocument.CreateElement('TestTable'); }

            Section -Name TestSection -ScriptBlock { Get-Process | Select-Object -First 3 | Table TestTable } | Out-XmlTable

            Assert-MockCalled -CommandName Out-XmlTable -Exactly 1
        }

        It 'calls Out-XmlParagraph' {
            Mock -CommandName Out-XmlParagraph -MockWith { return $xmlDocument.CreateElement('TestParagraph'); }

            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | Out-XmlParagraph

            Assert-MockCalled -CommandName Out-XmlParagraph -Exactly 1
        }

        It 'calls nested Out-XmlSection' {
            ## Note this must be called last in the Describe script block as the Out-XmlSection gets mocked!
            Mock -CommandName Out-XmlSection -Verifiable -MockWith { return $xmlDocument.CreateElement('TestSection'); }

            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | Out-XmlSection

            Assert-MockCalled -CommandName Out-XmlSection -Exactly 1
        }
    }
}

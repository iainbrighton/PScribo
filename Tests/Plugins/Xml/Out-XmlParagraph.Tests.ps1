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

    Describe 'Plugins\Xml\Out-XmlParagraph' {

        Context 'By Name Parameter.' {

            It 'outputs a XmlComment with no Text' {
                $text = 'Test paragraph.'
                $p = Paragraph -Name $text | Out-XmlParagraph
                $p.GetType() | Should Be 'System.Xml.XmlComment'
                $p.Value.Trim() | Should BeExactly $text
            }

            It 'outputs a XmlElement with a Name/Id' {
                $name = 'testid'
                $text = 'Test paragraph.'
                $p = Paragraph -Name $name -Text $text | Out-XmlParagraph
                $p.Name | Should BeExactly $name
                $p.'#text' | Should BeExactly $text
            }

            It 'outputs a XmlElement with a Name/Id and Value' {
                $name = 'testid'
                $text = 'Test paragraph.'
                $value = 'Customised test paragraph.'
                $p = Paragraph -Name $name -Text $text -Value $value | Out-XmlParagraph
                $p.Name | Should BeExactly $name
                $p.'#text' | Should BeExactly $value
            }

        } #end context By Name Parameter

        Context 'By Positional Parameter.' {

            It 'outputs a XmlComment with no Text' {
                $text = 'Test paragraph.'
                $p = Paragraph $text | Out-XmlParagraph
                $p.GetType() | Should Be 'System.Xml.XmlComment'
                $p.Value.Trim() | Should BeExactly $text
            }

             It 'outputs a XmlElement with a Name/Id' {
                $name = 'testid'
                $text = 'Test paragraph.'
                $p = Paragraph $name $text | Out-XmlParagraph
                $p.Name | Should BeExactly $name
                $p.'#text' | Should BeExactly $text
            }

            It 'outputs a XmlElement with a Name/Id and Value' {
                $name = 'testid'
                $text = 'Test paragraph.'
                $value = 'Customised test paragraph.'
                $p = Paragraph $name $text $value | Out-XmlParagraph
                $p.Name | Should BeExactly $name
                $p.'#text' | Should BeExactly $value
            }
        }
    }
}

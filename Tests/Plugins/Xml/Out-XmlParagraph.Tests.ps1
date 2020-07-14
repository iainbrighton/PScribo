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

    Describe 'Plugins\Xml\Out-XmlParagraph' {

        ## Scaffold document
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {
            Paragraph -Name 'TestParagraph';
            Get-Process | Select-Object -First 3 | Table 'TestTable';
            Section -Name 'TestSection' -ScriptBlock {
                Section -Name 'TestSubSection' { }
            }
        }
        $Document = $pscriboDocument
        $script:currentPageNumber = 1

        Context 'By Name Parameter.' {

            It 'outputs a named "paragraph" XmlElement' {
                $name = 'testid'

                $p = Paragraph -Name $name | Out-XmlParagraph

                $p.GetType() | Should Be 'System.Xml.XmlElement'
                $p.Name | Should BeExactly 'paragraph'
            }

            It 'outputs textnode with "Name" contents' {
                $name = 'testid'

                $p = Paragraph -Name $name | Out-XmlParagraph

                $p.'#text' | Should BeExactly $name
            }

            It 'outputs textnode with "Text" contents' {
                $name = 'testid'
                $text = 'Test paragraph'

                $p = Paragraph -Name $name -Text $text | Out-XmlParagraph

                $p.'#text' | Should BeExactly $text
            }

            It 'outputs space between text runs (by default)' {
                $paragraph = {
                    Text 'Test'
                    Text 'paragraph'
                }

                $result = Paragraph -ScriptBlock $paragraph | Out-XmlParagraph

                $result.'#text' | Should BeExactly 'Test paragraph'
            }

            It 'does not output space between text runs (when specified)' {
                $paragraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }

                $result = Paragraph -ScriptBlock $paragraph | Out-XmlParagraph

                $result.'#text' | Should BeExactly 'Testparagraph'
            }

        } #end context By Name Parameter

        Context 'By Positional Parameter.' {

            It 'outputs a named "paragraph" XmlElement' {
                $name = 'testid'

                $p = Paragraph $name | Out-XmlParagraph

                $p.GetType() | Should Be 'System.Xml.XmlElement'
                $p.Name | Should BeExactly 'paragraph'
            }

            It 'outputs textnode with "Name" contents' {
                $name = 'testid'

                $p = Paragraph $name | Out-XmlParagraph

                $p.'#text' | Should BeExactly $name
            }

            It 'outputs textnode with "Text" contents' {
                $name = 'testid'
                $text = 'Test paragraph'

                $p = Paragraph $name $text | Out-XmlParagraph

                $p.'#text' | Should BeExactly $text
            }

            It 'outputs space between text runs (by default)' {
                $paragraph = {
                    Text 'Test'
                    Text 'paragraph'
                }

                $result = Paragraph $paragraph | Out-XmlParagraph

                $result.'#text' | Should BeExactly 'Test paragraph'
            }

            It 'does not output space between text runs (when specified)' {
                $paragraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }

                $result = Paragraph $paragraph | Out-XmlParagraph

                $result.'#text' | Should BeExactly 'Testparagraph'
            }
        }
    }
}

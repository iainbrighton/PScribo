$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$moduleRoot = Split-Path -Path $here -Parent;
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
        Get-Service | Select-Object -First 3 | Table 'TestTable';
        Section -Name 'TestSection' -ScriptBlock {
            Section -Name 'TestSubSection' { }
        }
    }

    Describe 'OutXml' {
        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'calls OutXmlTable' {
            Mock -CommandName OutXmlTable -MockWith { return $xmlDocument.CreateElement('TestTable') };
            Document -Name 'TestDocument' -ScriptBlock { Get-Service | Select-Object -First 1 | Table 'TestTable' } | OutXml -Path $path;
            Assert-MockCalled -CommandName OutXmlTable -Exactly 1;
        }
    
        It 'calls OutXmlParagraph' {
            Mock -CommandName OutXmlParagraph -MockWith { return $xmlDocument.CreateElement('TestParagraph') };
            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | OutXml -Path $path;
            Assert-MockCalled -CommandName OutXmlParagraph -Exactly 1;
        }
    
        It 'calls OutXmlSection' {
            Mock -CommandName OutXmlSection -MockWith { return $xmlDocument.CreateElement('TestSection'); };
            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | OutXml -Path $path;
            Assert-MockCalled -CommandName OutXmlSection -Exactly 1;
        }

    } #end describe OutXml

    Describe 'OutXmlParagraph' {
    
        Context 'By Name Parameter.' {

            It 'outputs a XmlComment with no Text' {
                $text = 'Test paragraph.';
                $p = Paragraph -Name $text | OutXmlParagraph;
                $p.GetType() | Should Be 'System.Xml.XmlComment';
                $p.Value.Trim() | Should BeExactly $text;
            }

            It 'outputs a XmlElement with a Name/Id' {
                $name = 'testid';
                $text = 'Test paragraph.';
                $p = Paragraph -Name $name -Text $text | OutXmlParagraph;
                $p.Name | Should BeExactly $name;
                $p.'#text' | Should BeExactly $text;
            }

            It 'outputs a XmlElement with a Name/Id and Value' {
                $name = 'testid';
                $text = 'Test paragraph.';
                $value = 'Customised test paragraph.';
                $p = Paragraph -Name $name -Text $text -Value $value | OutXmlParagraph;
                $p.Name | Should BeExactly $name;
                $p.'#text' | Should BeExactly $value;
            }
    
        } #end context By Name Parameter

        Context 'By Positional Parameter.' {
            It 'outputs a XmlComment with no Text' {
                $text = 'Test paragraph.';
                $p = Paragraph $text | OutXmlParagraph;
                $p.GetType() | Should Be 'System.Xml.XmlComment';
                $p.Value.Trim() | Should BeExactly $text;
            }

             It 'outputs a XmlElement with a Name/Id' {
                $name = 'testid';
                $text = 'Test paragraph.';
                $p = Paragraph $name $text | OutXmlParagraph;
                $p.Name | Should BeExactly $name;
                $p.'#text' | Should BeExactly $text;
            }

            It 'outputs a XmlElement with a Name/Id and Value' {
                $name = 'testid';
                $text = 'Test paragraph.';
                $value = 'Customised test paragraph.';
                $p = Paragraph $name $text $value | OutXmlParagraph;
                $p.Name | Should BeExactly $name;
                $p.'#text' | Should BeExactly $value;
            }
        } #end context By Positional Parameter

    } #end describe OutXmlParagraph

    Describe 'OutXmlSection' {

        It 'calls OutXmlTable' {
            Mock -CommandName OutXmlTable -MockWith { return $xmlDocument.CreateElement('TestTable'); };
            Section -Name TestSection -ScriptBlock { Get-Service | Select-Object -First 3 | Table TestTable } | OutXmlTable;
            Assert-MockCalled -CommandName OutXmlTable -Exactly 1;
        }

        It 'calls OutXmlParagraph' {
            Mock -CommandName OutXmlParagraph -MockWith { return $xmlDocument.CreateElement('TestParagraph'); };
            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | OutXmlParagraph;
            Assert-MockCalled -CommandName OutXmlParagraph -Exactly 1;
        }

        It 'calls nested OutXmlSection' {
            ## Note this must be called last in the Describe script block as the OutXmlSection gets mocked!
            Mock -CommandName OutXmlSection -Verifiable -MockWith { return $xmlDocument.CreateElement('TestSection'); };
            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | OutXmlSection;
            Assert-MockCalled -CommandName OutXmlSection -Exactly 1;
        }

    } #end describe OutXmlSection

    Describe 'OutXmlTable' {
        $services = Get-Service | Select -First 3;
        $tableName = 'Test Table';
        $tableColumns = @('Name','DisplayName','Status');
        $tableHeaders = @('Name','Display Name','Status');

        Context 'By Name Parameter.' {
        
            It 'outputs a root XmlElement' {
                $table = $services | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    OutXmlTable;
                $table.LocalName | Should BeExactly ($tableName.Replace(' ','').ToLower());
                $table.Name | Should BeExactly $tableName;

            }
    
            It 'outputs a XmlElement for each table row' {
                $table = $services | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    OutXmlTable;
                $table.ChildNodes.Count | Should Be $services.Count;
            }

            It 'outputs a XmlElement for each table column' {
                $table = $services | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    OutXmlTable;
                $table.FirstChild.ChildNodes.Count | Should Be $tableColumns.Count;
            }

            It 'creates a name attribute when headers contain spaces' {
                $table = $services | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    OutXmlTable;
                for ($i = 0; $i -lt $table.FirstChild.ChildNodes.Count; $i++) { 
                    $table.FirstChild.ChildNodes[$i].LocalName | Should BeExactly $tableColumns[$i].ToLower();
                    if ($tableHeaders[$i].Contains(' ')) {
                        $table.FirstChild.ChildNodes[$i].Name | Should BeExactly $tableHeaders[$i];
                    }
                }
            }
        } #end context By Name Parameter

         Context 'By Postional Parameter.' {
        
            It 'outputs a root XmlElement' {
                $table = $services | Table $tableName $tableColumns $tableHeaders |
                    OutXmlTable;
                $table.LocalName | Should BeExactly ($tableName.Replace(' ','').ToLower());
                $table.Name | Should BeExactly $tableName;
                $table.ChildNodes.Count | Should Be 3;
            }
    
            It 'outputs a XmlElement for each table row' {
                $table = $services | Table $tableName $tableColumns $tableHeaders |
                    OutXmlTable;
                $table.ChildNodes.Count | Should Be $services.Count;
            }

            It 'outputs a XmlElement for each table column' {
                $table = $services | Table $tableName $tableColumns $tableHeaders |
                    OutXmlTable;
                $table.FirstChild.ChildNodes.Count | Should Be $tableColumns.Count;
            }

            It 'creates a name attribute when headers contain spaces' {
                $table = $services | Table $tableName $tableColumns $tableHeaders |
                    OutXmlTable;
                for ($i = 0; $i -lt $table.FirstChild.ChildNodes.Count; $i++) { 
                    $table.FirstChild.ChildNodes[$i].LocalName | Should BeExactly $tableColumns[$i].ToLower();
                    if ($tableHeaders[$i].Contains(' ')) {
                        $table.FirstChild.ChildNodes[$i].Name | Should BeExactly $tableHeaders[$i];
                    }
                }
            }
        } #end context By Positional Parameter

    } #end describe OutXmlTable

} #end inmodulescope 
 
<#
Code coverage report:
Covered 59.62% of 52 analyzed commands in 1 file.

Missed commands:

File                Function      Line Command                                                                                                                      
----                --------      ---- -------                                                                                                                      
OutXml.Internal.ps1 OutXmlSection   13 $sectionId = ($Section.Id -replace '[^a-z0-9-_\.]','').ToLower()                                                             
OutXml.Internal.ps1 OutXmlSection   13 $Section.Id -replace '[^a-z0-9-_\.]',''                                                                                      
OutXml.Internal.ps1 OutXmlSection   14 $element = $xmlDocument.CreateElement($sectionId)                                                                            
OutXml.Internal.ps1 OutXmlSection   15 [ref] $null = $element.SetAttribute("name", $Section.Name)                                                                   
OutXml.Internal.ps1 OutXmlSection   16 $Section.Sections.GetEnumerator()                                                                                            
OutXml.Internal.ps1 OutXmlSection   17 if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }...                                               
OutXml.Internal.ps1 OutXmlSection   17 $sectionId = '{0}..' -f $s.Id.Substring(0,38)                                                                                
OutXml.Internal.ps1 OutXmlSection   18 $sectionId = $s.Id                                                                                                           
OutXml.Internal.ps1 OutXmlSection   19 WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent ($s.Level +1)                          
OutXml.Internal.ps1 OutXmlSection   19 $localized.PluginProcessingSection -f $s.Type, $sectionId                                                                    
OutXml.Internal.ps1 OutXmlSection   19 $s.Level +1                                                                                                                  
OutXml.Internal.ps1 OutXmlSection   20 $s.Type                                                                                                                      
OutXml.Internal.ps1 OutXmlSection   21 [ref] $null = $element.AppendChild((OutXmlSection -Section $s))                                                              
OutXml.Internal.ps1 OutXmlSection   21 OutXmlSection -Section $s                                                                                                    
OutXml.Internal.ps1 OutXmlSection   22 [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s))                                                          
OutXml.Internal.ps1 OutXmlSection   22 OutXmlParagraph -Paragraph $s                                                                                                
OutXml.Internal.ps1 OutXmlSection   23 [ref] $null = $element.AppendChild((OutXmlTable -Table $s))                                                                  
OutXml.Internal.ps1 OutXmlSection   23 OutXmlTable -Table $s                                                                                                        
OutXml.Internal.ps1 OutXmlSection   29 WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning                                                
OutXml.Internal.ps1 OutXmlSection   29 $localized.PluginUnsupportedSection -f $s.Type                                                                               
OutXml.Internal.ps1 OutXmlSection   33 return $element   
#>
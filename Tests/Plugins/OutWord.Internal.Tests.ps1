$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {
    
    function NewTestDocument {
        [CmdletBinding()]
        param (
            [Switch] $NoTestElement
        )
        $testNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $newTestDocument = New-Object -TypeName 'System.Xml.XmlDocument';
        [ref] $null = $newTestDocument.AppendChild($newTestDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
        if (-not $NoTestElement) {
            [ref] $null = $newTestDocument.AppendChild($newTestDocument.CreateElement('w', 'test', $testNamespace));
        }
        return $newTestDocument;
    }
    
    function GetMatch {
        [CmdletBinding()]
        param (
            [System.String] $String,
            
            [System.Management.Automation.SwitchParameter] $Complete
        )
        Write-Verbose "Pre Match : '$String'";
        $matchString = $String.Replace('/','\/');
        if (-not $String.StartsWith('^')) {
            $matchString = $matchString.Replace('[..]','[\s\S]+');
            if ($Complete) {
                $matchString = '^<w:test xmlns:w="http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main">{0}<\/w:test>$' -f $matchString; 
            }
        }
        Write-Verbose "Post Match: '$matchString'";
        return $matchString;
    } #end function GetMatch
    
    Describe 'OutWord.Internal\ConvertToWordColor' {
    
        It 'converts to "abcdef" to "ABCDEF"' {
            $result = ConvertToWordColor 'abcdef';
            
            $result | Should BeExactly 'ABCDEF';
        }
        
        It 'converts "#abcdef" to "ABCDEF"' {
            $result = ConvertToWordColor '#abcdef';
            
            $result | Should BeExactly 'ABCDEF';
        }
        
        It 'converts "abc" to "AABBCC"' {
            $result = ConvertToWordColor 'abc';
            
            $result | Should BeExactly 'AABBCC';
        }
        
        It 'converts "#abc" to "AABBCC"' {
            $result = ConvertToWordColor '#abc';
            
            $result | Should BeExactly 'AABBCC';
        }
        
    } #end describe OutWord.Internal\ConvertToWordColor
    
    Describe 'OutWord.Internal\OutWordSection' {
        
            It 'appends section "<w:p>[..]</w:p>"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name TestSection -ScriptBlock { };
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '<w:p>[..]</w:p>';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            It 'appends section spacing "[..]<w:pPr><w:spacing w:before="160" w:after="160" /></w:pPr>[..]"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name TestSection -ScriptBlock { };
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '[..]<w:pPr><w:spacing w:before="160" w:after="160" /></w:pPr>[..]';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            
            It 'appends section style "[..]<w:pStyle w:val="CustomStyle" />[..]"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name TestSection -ScriptBlock { };
                $testSection.Style = 'CustomStyle';
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '[..]<w:pStyle w:val="CustomStyle" />[..]';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            
            It 'increases section spacing between section levels' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name TestSection -ScriptBlock { };
                $testSection.Level = 3;
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '[..]<w:pPr><w:spacing w:before="280" w:after="280" /></w:pPr>[..]';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            
            It 'appends "[..]<w:r><w:t>Section Run</w:t></w:r></w:p>" run' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name 'Section Run' -ScriptBlock { };
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '[..]<w:r><w:t>Section Run</w:t></w:r></w:p>';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            
            It 'adds section numbering when enabled' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $Document.Options['EnableSectionNumbering'] = $true;
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                $testSection = Section -Name 'Numbered Section' -ScriptBlock { };
                $testSection.Number = 2;
                
                OutWordSection -Section $testSection -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                $expected = GetMatch '[..]<w:r><w:t>2 Numbered Section</w:t></w:r></w:p>';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }
            
            It 'calls "OutWordParagraph"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordParagraph { return $testDocument.CreateElement('mockParagraph'); };
                
                $section = Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' }; 
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordParagraph -Scope It;
            }

            It 'calls "OutTextParagraph" twice' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordParagraph { return $testDocument.CreateElement('mockParagraph'); };
                
                $section = Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph'; Paragraph 'TestParagraph'; }; 
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordParagraph -Exactly 2 -Scope It;
            }

            It 'calls "OutWordTable"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordTable { return $testDocument.CreateElement('mockTable'); };

                $section = Section -Name TestSection -ScriptBlock { Get-Service | Select-Object -First 3 | Table TestTable };
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;

                Assert-MockCalled -CommandName OutWordTable -Scope It;
            }

            It 'calls "OutWordPageBreak"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordPageBreak { return $testDocument.CreateElement('mockPageBreak'); };

                $section = Section -Name TestSection -ScriptBlock { PageBreak };
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordPageBreak -Scope It;
            }

            It 'calls "OutWordLineBreak"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordLineBreak { return $testDocument.CreateElement('mockLineBreak'); };

                $section = Section -Name TestSection -ScriptBlock { LineBreak };
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordLineBreak -Scope It;
            }

            It 'calls "OutTextBlankLine"' {
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordBlankLine { return $testDocument.CreateElement('mockBlankLine'); };

                $section = Section -Name TestSection -ScriptBlock { BlankLine };
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordBlankLine -Scope It;
            }

            It 'calls nested "OutWordSection"' {
                ## Note this must be called last in the Describe script block as the OutXmlSection gets mocked!
                $Document = Document -Name 'TestDocument' -ScriptBlock { };
                $pscriboDocument = $Document;
                $testDocument = NewTestDocument;
                Mock OutWordSection -MockWith { return $testDocument.CreateElement('mockSection'); };
                
                $section = Section -Name TestSection -ScriptBlock { Section -Name SubSection { } };
                OutWordSection -Section $section -XmlDocument $testDocument -RootElement $testDocument.DocumentElement;
                
                Assert-MockCalled -CommandName OutWordSection -Scope It;
            }
                
    } #end describe OutWord.Internal\OutWordSection
    
    Describe 'OutWord.Internal\OutWordParagraph' {
            
        It 'returns a "System.Xml.XmlElement" object type' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $result = OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument;
            
            $result -is [System.Xml.XmlElement] | Should Be $true;
        }
        
        It 'outputs paragraph "<w:p>[..]></w:p>"' {
            # creates paragraph property
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '<w:p>[..]></w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs paragraph properties "<w:p><w:pPr>[..]></w:pPr[..]></w:p>"' {
            # creates paragraph property
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '<w:p><w:pPr>[..]></w:pPr[..]></w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run "[..]<w:r>[..]></w:r>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:r>[..]></w:r>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs empty run properties "[..]<w:r><w:rPr />[..]></w:r>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:r><w:rPr />[..]></w:r>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property font "[..]<w:rPr><w:rFonts w:ascii="[..]" w:hAnsi="[..]" /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Font Ariel;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:rFonts w:ascii="Ariel" w:hAnsi="Ariel" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property font size "[..]<w:rPr><w:sz w:val="[..]" /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Size 10;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:sz w:val="20" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property bold "[..]<w:rPr><w:b /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Bold;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:b /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property italic "[..]<w:rPr><w:i /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Italic;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:i /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property underline "[..]<w:rPr><w:u w:val="single" /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Underline;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:u w:val="single" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run property colour "[..]<w:rPr><w:color w:val="112233" /></w:rPr>[..]"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Color 123;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:rPr><w:color w:val="112233" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs run text "[..]<w:r>[..]<w:t[..]</w:t>[..]"' {
            ## Ignore the space preservation namespace
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph';

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));
            
            $expected = GetMatch '[..]<w:r>[..]<w:t[..]</w:t>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected; 
        }
        
    } #end describe OutWord.Internal\OutWordParagraph

    Describe 'OutWord.Internal\OutWordPageBreak' {
        
        It 'returns a "System.Xml.XmlElement" object type' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument -NoTestElement;
            $testPageBreak = PageBreak;

            $result = OutWordPageBreak -PageBreak $testPageBreak -XmlDocument $testDocument;
            
            $result -is [System.Xml.XmlElement] | Should Be $true;
        }
        
        It 'outputs "<w:p><w:r><w:br w:type="page" /></w:r></w:p>"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument -NoTestElement;
            $testPageBreak = PageBreak;

            $result = OutWordPageBreak -PageBreak $testPageBreak -XmlDocument $testDocument;
        
            $expected = GetMatch '^<w:p [\s\S]+><w:r><w:br w:type="page" /></w:r></w:p>$'
            $result.OuterXml  | Should Match $expected; 
        }
        
    } #end describe OutWord.Internal\OutWordPageBreak
        
    Describe 'OutWord.Internal\OutWordLineBreak' {
        
        It 'returns a "System.Xml.XmlElement" object type' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument -NoTestElement;
            $testLineBreak = LineBreak;

            $result = OutWordLineBreak -LineBreak $testLineBreak -XmlDocument $testDocument;
            
            $result -is [System.Xml.XmlElement] | Should Be $true;
        }
        
        It 'outputs paragraph properties "<w:p><w:pPr>[..]</w:pPr></w:p>"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument -NoTestElement;
            $testLineBreak = LineBreak;

            $result = OutWordLineBreak -LineBreak $testLineBreak -XmlDocument $testDocument;
        
            $expected = GetMatch '^<w:p [\s\S]+><w:pPr>[\s\S]+</w:pPr></w:p>$'
            $result.OuterXml  | Should Match $expected; 
        }
        
        It 'outputs border "<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto" /></w:pBdr>"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument -NoTestElement;
            $testLineBreak = LineBreak;

            $result = OutWordLineBreak -LineBreak $testLineBreak -XmlDocument $testDocument;
        
            $expected = GetMatch '^<w:p [\s\S]+<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto" /></w:pBdr>[\s\S]+$'
            $result.OuterXml  | Should Match $expected; 
        }
        
    } #end describe OutWord.Internal\context OutWordLineBreak
        
    Describe 'OutWord.Internal\OutWordTable' {
        
        <#
        <w:test xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:tbl>
    <w:tr>
        <w:trPr>
            <w:tblHeader />
        </w:trPr>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:color="auto" w:fill="4472C4" />
                <w:tcW w:w="0" w:type="auto" />
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultHeading" />
                </w:pPr>
                <w:r>
                    <w:t>Name</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:color="auto" w:fill="4472C4" />
                <w:tcW w:w="0" w:type="auto" />
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultHeading" />
                </w:pPr>
                <w:r>
                    <w:t>DisplayName</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr>
        <w:tc>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultRow" />
                </w:pPr>
                <w:r>
                    <w:t>AJRouter</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultRow" />
                </w:pPr>
                <w:r>
                    <w:t>AllJoyn Router Service</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:color="auto" w:fill="D0DDEE" />
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultAltRow" />
                </w:pPr>
                <w:r>
                    <w:t>ALG</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:tcPr>
                <w:shd w:val="clear" w:color="auto" w:fill="D0DDEE" />
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultAltRow" />
                </w:pPr>
                <w:r>
                    <w:t>Application Layer Gateway Service</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
    <w:tr>
        <w:tc>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultRow" />
                </w:pPr>
                <w:r>
                    <w:t>AppIDSvc</w:t>
                </w:r>
            </w:p>
        </w:tc>
        <w:tc>
            <w:p>
                <w:pPr>
                    <w:pStyle w:val="TableDefaultRow" />
                </w:pPr>
                <w:r>
                    <w:t>Application Identity</w:t>
                </w:r>
            </w:p>
        </w:tc>
    </w:tr>
</w:tbl>
</w:test>
        #>
        BeforeEach {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $Document = $pscriboDocument;
            $testDocument = NewTestDocument;
            $testTable = Get-Service | Select -First 3 | Table -Name 'Test Table' -Columns 'Name','DisplayName';   
        }
        
        It 'appends table "<w:tbl>[..]</w:tbl>"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '<w:tbl>[..]</w:tbl>';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs table rows including header "(<w:tr>[..]?</w:tr>){4}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:tr>[..]?</w:tr>){4}'
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs table header "<w:tr><w:trPr><w:tblHeader />"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '<w:tr><w:trPr><w:tblHeader />';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs table borders "<w:tblGrid>[..]</w:tblGrid>"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '<w:tblGrid>[..]</w:tblGrid>';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs grid per column "(<w:gridCol />){2}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:gridCol />){2}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs table cells per row "(<w:tc>[..]?<\/w:tc>.*){8}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:tc>[..]?</w:tc>.*){8}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs paragraph per table cell "(<w:p>[..]</w:p>.*){8}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:p>[..]</w:p>.*){8}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs paragraph run style "(<w:p><w:pPr><w:pStyle w:val="[..]" /></w:pPr>[..]</w:p>.*){8}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:p><w:pPr><w:pStyle w:val="[..]" /></w:pPr>[..]</w:p>.*){8}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs paragraph run per cell "(<w:p>[..]<w:r><w:t>[..]</w:t></w:r></w:p>.*){8}"' {
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:p>[..]<w:r><w:t>[..]</w:t></w:r></w:p>.*){8}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs default table heading style "(<w:pStyle w:val="TableDefaultHeading" />.*){2}"' {
            ## 2 x heading cells
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:pStyle w:val="TableDefaultHeading" />.*){2}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs default table row style "(<w:pStyle w:val="TableDefaultRow" />.*){4}"' {
            ## 4 x default table row cells
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:pStyle w:val="TableDefaultRow" />.*){4}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'outputs alternate table row style "(<w:pStyle w:val="TableDefaultAltRow" />.*){2}"' {
            ## 2 x alternating table row cells
            OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '(<w:pStyle w:val="TableDefaultAltRow" />.*){2}';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
    } #end describe OutWord.Internal\OutWordTable

    Describe 'OutWord.Internal\OutWordBlankLine' {
        
        It 'appends paragraph "<w:p />"' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testBlankLine = BlankLine;

            OutWordBlankLine -BlankLine $testBlankLine -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '<w:p />';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
        It 'appends paragraph "<w:p />" per blankline' {
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testBlankLines = BlankLine -Count 2;

            OutWordBlankLine -BlankLine $testBlankLines -XmlDocument $testDocument -Element $testDocument.DocumentElement;
            
            $expected = GetMatch '<w:p /><w:p />';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
        
    } #end describe OutWord.Internal\OutWordBlankLine
    
} #end in module scope

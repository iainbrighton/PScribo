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

        It 'outputs indented paragraph "<w:p><w:pPr><w:ind w:left="1440" />[..]</w:p>"' {
            # creates paragraph property
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Tabs 2;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));

            $expected = GetMatch '<w:p><w:pPr><w:ind w:left="1440" />[..]</w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs paragraph style "<w:p><w:pPr><w:pStyle w:val="[..]" />[..]</w:p>"' {
            # creates paragraph property
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraph = Paragraph 'Test paragraph' -Style Heading3;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));

            $expected = GetMatch '<w:p><w:pPr><w:pStyle w:val="Heading3" />[..]</w:p>'
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

        It 'outputs run text "[..]<w:r>[..]<w:t [..]>{0}</w:t>[..]" using "Name" property' {
            ## Ignore the space preservation namespace
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraphText = 'Test paragraph';
            $testParagraph = Paragraph -Name $testParagraphText;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));

            $expected = GetMatch ('[..]<w:r>[..]<w:t [..]>{0}</w:t>[..]' -f $testParagraphText);
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run text "[..]<w:r>[..]<w:t [..]>{0}</w:t>[..]" using "Text" property' {
            ## Ignore the space preservation namespace
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraphText = 'Test paragraph';
            $testParagraph = Paragraph -Name 'Test' -Text $testParagraphText;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));

            $expected = GetMatch ('[..]<w:r>[..]<w:t [..]>{0}</w:t>[..]' -f $testParagraphText);
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run text "[..]<w:r>[..]<w:t [..]><w:br /></w:t>[..]</w:r>[..]" with embedded new line' {
            ## Ignore the space preservation namespace
            $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
            $testDocument = NewTestDocument;
            $testParagraphText = "Test`r`nParagraph";
            $testParagraph = Paragraph -Name 'Test' -Text $testParagraphText;

            $testDocument.DocumentElement.AppendChild((OutWordParagraph -Paragraph $testParagraph -XmlDocument $testDocument));

            $expected = GetMatch ('[..]<w:r>[..]<w:t [..]><w:br /></w:t>[..]</w:r>[..]');
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

        Context 'Default' {

            BeforeEach {
                $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
                $Document = $pscriboDocument;
                $testDocument = NewTestDocument;
            }

            foreach ($tableStyle in @($false, $true)) {

                $tableType = if ($tableStyle) { 'Tabular' } else { 'List' }

                It "outputs $tableType table border `"<w:tblBorders>[..]</w:tblBorders>`"" {
                    $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 3 | Table -Name 'Test Table' -List:$tableStyle;

                    OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                    $expected = GetMatch '<w:tblBorders>[..]</w:tblBorders>';
                    $testDocument.DocumentElement.OuterXml | Should Match $expected;
                }

                It "outputs $tableType table border color" {
                    $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 3 | Table -Name 'Test Table' -List:$tableStyle;
                    $borderColor = ($pscriboDocument.TableStyles[$testTable.Style]).BorderColor;

                    OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:top w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:bottom w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:start w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:end w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:insideH w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:insideV w:sz="5.76" w:val="single" w:color="{0}" />' -f $borderColor));
                }
                It "outputs $tableType table cell spacing `"<w:tblCellMar>[..]</w:tblCellMar>`"" {
                    $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 3 | Table -Name 'Test Table' -List:$tableStyle;
                    $paddingTop = ConvertToInvariantCultureString -Object (ConvertMmToTwips ($pscriboDocument.TableStyles[$testTable.Style]).PaddingTop);
                    $paddingLeft = ConvertToInvariantCultureString -Object (ConvertMmToTwips ($pscriboDocument.TableStyles[$testTable.Style]).PaddingLeft);
                    $paddingBottom = ConvertToInvariantCultureString -Object (ConvertMmToTwips ($pscriboDocument.TableStyles[$testTable.Style]).PaddingBottom);
                    $paddingRight = ConvertToInvariantCultureString -Object (ConvertMmToTwips ($pscriboDocument.TableStyles[$testTable.Style]).PaddingRight);

                    OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:tblCellMar>[..]</w:tblCellMar>' -f $paddingTop));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:top w:w="{0}" w:type="dxa" />' -f $paddingTop));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:start w:w="{0}" w:type="dxa" />' -f $paddingLeft));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:bottom w:w="{0}" w:type="dxa" />' -f $paddingBottom));
                    $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:end w:w="{0}" w:type="dxa" />' -f $paddingRight));
                }

                It "outputs $tableType table spacing `"<w:spacing w:before=`"72`" w:after=`"72`" />`"" {
                    $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 3 | Table -Name 'Test Table' -List:$tableStyle;
                    OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                    $expected = GetMatch ('<w:spacing w:before="72" w:after="72" />');
                    $testDocument.DocumentElement.OuterXml | Should Match $expected;
                }

            } #end foreach table type

        } #end context default

        Context 'List Table' {

            BeforeEach {
                $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
                $Document = $pscriboDocument;
                $testDocument = NewTestDocument;
                $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 3 | Table -Name 'Test Table' -List;
            }

            It 'outputs table per row "(<w:tbl>[..]</w:tbl>.*){3}"' {
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:tbl>[..]</w:tbl>.*){{{0}}}' -f ($testTable.Rows.Count));
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs space between each table "(<w:p />.*){2}"' {
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:p />.*){{{0}}}' -f ($testTable.Rows.Count -1));
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs one row per object property "(<w:tr>[..]</w:tr>.*){3}"' {
                $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 1 | Table -Name 'Test Table' -List;
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                ## Ignore __Style property
                $rowPropertyCount = @(($testTable.Rows[0]).PSObject.Properties).Count -1;
                $expected = GetMatch ('(<w:tr>[..]</w:tr>.*){{{0}}}' -f $rowPropertyCount);
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs two cells per object property "(<w:tc>[..]</w:tc>.*){6}"' {
                $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 1 | Table -Name 'Test Table' -List;
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                ## Ignore __Style property
                $rowPropertyCount = @(($testTable.Rows[0]).PSObject.Properties).Count -1;
                $expected = GetMatch ('(<w:tc>[..]</w:tc>.*){{{0}}}' -f ($rowPropertyCount * 2));
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs table cell percentage widths "(<w:tcW w:w="[..]" w:type="pct" />.*){6}' {
                $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 1 | Table -Name 'Test Table' -List -ColumnWidths 30,70;
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:tcW w:w="[..]" w:type="pct" />.*){6}')
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs paragraph per table cell "(<w:p>[..]</w:p>.*){6}"' {
                $testTable = Get-Service | Select -Property 'Name','DisplayName','Status' -First 1 | Table -Name 'Test Table' -List -ColumnWidths 30,70;
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch '(<w:p>[..]</w:p>.*){6}';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs custom cell style "(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>){1}"' {
                $testStyleName = 'Title';
                $testTable = @(
                    [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                    [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; 'Column2__Style' = $testStyleName; }
                )
                OutWordTable (Table 'TestTable' -Hashtable $testTable -List) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>){{1}}' -f $testStyleName);
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            ##It 'outputs custom row style "(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>.*){2}"' {
            ##    $testStyleName = 'Title';
            ##    $testTable = @(
            ##        [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
            ##        [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; '__Style' = $testStyleName; }
            ##    )
            ##    OutWordTable (Table 'TestTable' -Hashtable $testTable -List) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

            ##    $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>.*){{2}}' -f $testStyleName);
            ##    $testDocument.DocumentElement.OuterXml | Should Match $expected;
            ##}

            It 'outputs table cell with embedded new line' {
                $testTable = [Ordered] @{ Licenses = "Standard`r`nProfessional`r`nEnterprise"; }

                OutWordTable (Table 'TestTable' -Hashtable $testTable -List) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                ## Three lines = 2 line breaks
                $null = $testDocument.DocumentElement.OuterXml -match (GetMatch '(<w:r><w:t><w:br /></w:t></w:r>.*)');
                $matches.Count | Should Be 2;
            }

        } #end context list table

        Context 'Tabular Table' {

            BeforeEach {
                $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
                $Document = $pscriboDocument;
                $testDocument = NewTestDocument;
                $testTable = Get-Service | Select -First 3 | Table -Name 'Test Table' -Columns 'Name','DisplayName' -ColumnWidths 30,70;
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

            It 'outputs table cell percentage widths "(<w:tcW w:w="[..]" w:type="pct" />.*){2}"' {
                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch '(<w:tcW w:w="[..]" w:type="pct" />.*){2}';
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs custom cell style "(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>){1}"' {
                $testStyleName = 'Title';
                $testTable = @(
                    [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                    [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; 'Column2__Style' = $testStyleName; }
                )
                OutWordTable (Table 'TestTable' -Hashtable $testTable) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>){{1}}' -f $testStyleName);
                $testDocument.DocumentElement.OuterXml | Should Match $expected;
            }

            It 'outputs custom row style "(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>.*){2}"' {
                $testStyleName = 'Title';
                $testTable = @(
                    [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                    [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; '__Style' = $testStyleName; }
                )
                OutWordTable (Table 'TestTable' -Hashtable $testTable) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /></w:pPr>.*){{2}}' -f $testStyleName);
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

            It 'outputs custom table heading style "(<w:pStyle w:val="CustomStyle" />.*){2}"' {
                ## 2 x heading cells
                $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {
                    Style -Name 'CustomStyle' -Size 11 -Color 000;
                    TableStyle -Id 'CustomTableStyle' -HeaderStyle 'CustomStyle' -RowStyle 'CustomStyle' -AlternateRowStyle 'CustomStyle';
                };
                $Document = $pscriboDocument;
                $testDocument = NewTestDocument;
                $testTable = Get-Service | Select -First 3 | Table -Name 'Test Table' -Columns 'Name','DisplayName' -Style 'CustomTableStyle';

                OutWordTable $testTable -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $expected = GetMatch '(<w:pStyle w:val="CustomStyle" />.*){2}';
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

            It 'outputs table cell with embedded new line' {
                $licenses = "Standard`r`nProfessional`r`nEnterprise"
                $testTable = [Ordered] @{ Licenses = $licenses; }

                OutWordTable (Table 'TestTable' -Hashtable $testTable) -XmlDocument $testDocument -Element $testDocument.DocumentElement;

                $null = $testDocument.DocumentElement.OuterXml -match (GetMatch '(<w:r><w:t><w:br /></w:t></w:r>.*)');
                $matches.Count | Should Be 2;
            }

        } #end context Tabular Table

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

<#
Code coverage report:
Covered 51.53% of 522 analyzed commands in 1 file.

Missed commands:

File                 Function                Line Command
----                 --------                ---- -------
OutWord.Internal.ps1 OutWordSection            56 $sectionId = '{0}[..]' -f $s.Id.Substring(0,36)
OutWord.Internal.ps1 OutWordSection            59 $currentIndentationLevel = $s.Level +1
OutWord.Internal.ps1 OutWordSection            62 $s
OutWord.Internal.ps1 OutWordSection            62 OutWordSection -RootElement $RootElement -XmlDocument $XmlDocument
OutWord.Internal.ps1 OutWordSection            68 WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type)...
OutWord.Internal.ps1 OutWordSection            68 $localized.PluginUnsupportedSection -f $s.Type
OutWord.Internal.ps1 GetWordTable             205 $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblI...
OutWord.Internal.ps1 GetWordTable             206 [ref] $null = $tblInd.SetAttribute('w', $xmlnsMain, (720 * $Table....
OutWord.Internal.ps1 GetWordTable             206 720 * $Table.Tabs
OutWord.Internal.ps1 GetWordTable             213 $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 't...
OutWord.Internal.ps1 GetWordTable             214 [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'autofit')
OutWord.Internal.ps1 GetWordTable             224 $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options...
OutWord.Internal.ps1 GetWordTable             224 $Document.Options['PageMarginLeft'] + $Document.Options['PageMargi...
OutWord.Internal.ps1 GetWordTable             225 $indentWidthMm = ConvertPtToMm -Point ($Table.Tabs * 36)
OutWord.Internal.ps1 GetWordTable             225 $Table.Tabs * 36
OutWord.Internal.ps1 GetWordTable             226 $tableRenderMm = (($pageWidthMm / 100) * $Table.Width) + $indentWi...
OutWord.Internal.ps1 GetWordTable             226 ($pageWidthMm / 100) * $Table.Width
OutWord.Internal.ps1 GetWordTable             226 $pageWidthMm / 100
OutWord.Internal.ps1 GetWordTable             227 if ($tableRenderMm -gt $pageWidthMm) {...
OutWord.Internal.ps1 GetWordTable             229 $maxTableWidthMm = $pageWidthMm - $indentWidthMm
OutWord.Internal.ps1 GetWordTable             230 $tableWidthRenderPct = [System.Math]::Round(($maxTableWidthMm / $p...
OutWord.Internal.ps1 GetWordTable             230 $maxTableWidthMm / $pageWidthMm
OutWord.Internal.ps1 GetWordTable             231 WriteLog -Message ($localized.TableWidthOverflowWarning -f $tableW...
OutWord.Internal.ps1 GetWordTable             231 $localized.TableWidthOverflowWarning -f $tableWidthRenderPct
OutWord.Internal.ps1 OutWordTable             357 if (-not (Test-Path -Path Variable:\cellStyle)) {...
OutWord.Internal.ps1 OutWordTable             357 Test-Path -Path Variable:\cellStyle
OutWord.Internal.ps1 OutWordTable             358 $cellStyle = $Document.Styles[$row.$cellPropertyStyle]
OutWord.Internal.ps1 OutWordTable             360 if (-not (Test-Path -Path Variable:\cellStyle)) {...
OutWord.Internal.ps1 OutWordTable             362 $cellStyle = $Document.Styles[$row.$cellPropertyStyle]
OutWord.Internal.ps1 OutWordTable             364 if ($cellStyle.BackgroundColor) {...
OutWord.Internal.ps1 OutWordTable             365 [ref] $null = $tc2.AppendChild((GetWordTableStyleCellPr -Style $ce...
OutWord.Internal.ps1 OutWordTable             365 GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument
OutWord.Internal.ps1 OutWordTable             367 if ($row.$cellPropertyStyle) {...
OutWord.Internal.ps1 OutWordTable             368 $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xm...
OutWord.Internal.ps1 OutWordTable             369 $pStyle2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'pSty...
OutWord.Internal.ps1 OutWordTable             370 [ref] $null = $pStyle2.SetAttribute('val', $xmlnsMain, $row.$cellP...
OutWord.Internal.ps1 OutWordTOC               476 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 OutWordTOC               477 $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlnsMain)
OutWord.Internal.ps1 OutWordTOC               478 $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr',...
OutWord.Internal.ps1 OutWordTOC               479 $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', '...
OutWord.Internal.ps1 OutWordTOC               480 $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateEle...
OutWord.Internal.ps1 OutWordTOC               481 [ref] $null = $docObjectGallery.SetAttribute('val', $xmlnsMain, 'T...
OutWord.Internal.ps1 OutWordTOC               482 [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('...
OutWord.Internal.ps1 OutWordTOC               483 $sdtEndPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdE...
OutWord.Internal.ps1 OutWordTOC               485 $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'st...
OutWord.Internal.ps1 OutWordTOC               486 $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p',...
OutWord.Internal.ps1 OutWordTOC               487 $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xm...
OutWord.Internal.ps1 OutWordTOC               488 $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pSty...
OutWord.Internal.ps1 OutWordTOC               489 [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, 'TOC')
OutWord.Internal.ps1 OutWordTOC               490 $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               491 $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               492 [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name))
OutWord.Internal.ps1 OutWordTOC               494 $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p',...
OutWord.Internal.ps1 OutWordTOC               495 $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xm...
OutWord.Internal.ps1 OutWordTOC               496 $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs',...
OutWord.Internal.ps1 OutWordTOC               497 $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', ...
OutWord.Internal.ps1 OutWordTOC               498 [ref] $null = $tab2.SetAttribute('val', $xmlnsMain, 'right')
OutWord.Internal.ps1 OutWordTOC               499 [ref] $null = $tab2.SetAttribute('leader', $xmlnsMain, 'dot')
OutWord.Internal.ps1 OutWordTOC               500 [ref] $null = $tab2.SetAttribute('pos', $xmlnsMain, '9016')
OutWord.Internal.ps1 OutWordTOC               501 $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               503 $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldCh...
OutWord.Internal.ps1 OutWordTOC               504 [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlnsMain, 'b...
OutWord.Internal.ps1 OutWordTOC               506 $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               507 $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'inst...
OutWord.Internal.ps1 OutWordTOC               508 [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/...
OutWord.Internal.ps1 OutWordTOC               509 [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode('...
OutWord.Internal.ps1 OutWordTOC               511 $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               512 $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldCh...
OutWord.Internal.ps1 OutWordTOC               513 [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlnsMain, 's...
OutWord.Internal.ps1 OutWordTOC               515 $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p',...
OutWord.Internal.ps1 OutWordTOC               516 $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsM...
OutWord.Internal.ps1 OutWordTOC               518 $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldCh...
OutWord.Internal.ps1 OutWordTOC               519 [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlnsMain, 'e...
OutWord.Internal.ps1 OutWordTOC               521 return $sdt
OutWord.Internal.ps1 GetWordStyle             558 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 GetWordStyle             559 if ($Type -eq 'Paragraph') {...
OutWord.Internal.ps1 GetWordStyle             560 $styleId = $Style.Id
OutWord.Internal.ps1 GetWordStyle             561 $styleName = $Style.Name
OutWord.Internal.ps1 GetWordStyle             562 $linkId = '{0}Char' -f $Style.Id
OutWord.Internal.ps1 GetWordStyle             565 $styleId = '{0}Char' -f $Style.Id
OutWord.Internal.ps1 GetWordStyle             566 $styleName = '{0} Char' -f $Style.Name
OutWord.Internal.ps1 GetWordStyle             567 $linkId = $Style.Id
OutWord.Internal.ps1 GetWordStyle             571 $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
OutWord.Internal.ps1 GetWordStyle             572 [ref] $null = $documentStyle.SetAttribute('type', $xmlnsMain, $Typ...
OutWord.Internal.ps1 GetWordStyle             573 if ($Style.Id -eq $Document.DefaultStyle) {...
OutWord.Internal.ps1 GetWordStyle             575 [ref] $null = $documentStyle.SetAttribute('default', $xmlnsMain, 1)
OutWord.Internal.ps1 GetWordStyle             576 $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElemen...
OutWord.Internal.ps1 GetWordStyle             577 [ref] $null = $uiPriority.SetAttribute('val', $xmlnsMain, 1)
OutWord.Internal.ps1 GetWordStyle             579 if ($Style.Id -eq $Document.DefaultStyle) {...
OutWord.Internal.ps1 GetWordStyle             579 $Style.Id -eq 'Footer'
OutWord.Internal.ps1 GetWordStyle             579 $Style.Id -eq 'Header'
OutWord.Internal.ps1 GetWordStyle             581 [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElemen...
OutWord.Internal.ps1 GetWordStyle             583 if ($Style.Id -eq $Document.DefaultStyle) {...
OutWord.Internal.ps1 GetWordStyle             583 $document.TableStyles.Values
OutWord.Internal.ps1 GetWordStyle             583 ForEach-Object { $_.HeaderStyle; $_.RowStyle; $_.AlternateRowStyle; }
OutWord.Internal.ps1 GetWordStyle             583 $_.HeaderStyle
OutWord.Internal.ps1 GetWordStyle             583 $_.RowStyle
OutWord.Internal.ps1 GetWordStyle             583 $_.AlternateRowStyle
OutWord.Internal.ps1 GetWordStyle             585 [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElemen...
OutWord.Internal.ps1 GetWordStyle             588 [ref] $null = $documentStyle.SetAttribute('styleId', $xmlnsMain, $...
OutWord.Internal.ps1 GetWordStyle             589 $documentStyleName = $documentStyle.AppendChild($xmlDocument.Creat...
OutWord.Internal.ps1 GetWordStyle             590 [ref] $null = $documentStyleName.SetAttribute('val', $xmlnsMain, $...
OutWord.Internal.ps1 GetWordStyle             591 $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('...
OutWord.Internal.ps1 GetWordStyle             592 [ref] $null = $basedOn.SetAttribute('val', $XmlnsMain, 'Normal')
OutWord.Internal.ps1 GetWordStyle             593 $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w',...
OutWord.Internal.ps1 GetWordStyle             594 [ref] $null = $link.SetAttribute('val', $XmlnsMain, $linkId)
OutWord.Internal.ps1 GetWordStyle             595 $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w',...
OutWord.Internal.ps1 GetWordStyle             596 [ref] $null = $next.SetAttribute('val', $xmlnsMain, 'Normal')
OutWord.Internal.ps1 GetWordStyle             597 $qFormat = $documentStyle.AppendChild($XmlDocument.CreateElement('...
OutWord.Internal.ps1 GetWordStyle             598 $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', ...
OutWord.Internal.ps1 GetWordStyle             599 $keepNext = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keep...
OutWord.Internal.ps1 GetWordStyle             600 $keepLines = $pPr.AppendChild($XmlDocument.CreateElement('w', 'kee...
OutWord.Internal.ps1 GetWordStyle             601 $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spaci...
OutWord.Internal.ps1 GetWordStyle             602 [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
OutWord.Internal.ps1 GetWordStyle             603 [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
OutWord.Internal.ps1 GetWordStyle             605 $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmln...
OutWord.Internal.ps1 GetWordStyle             606 if ($Style.Align.ToLower() -eq 'justify') {...
OutWord.Internal.ps1 GetWordStyle             607 [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
OutWord.Internal.ps1 GetWordStyle             610 [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToL...
OutWord.Internal.ps1 GetWordStyle             612 if ($Style.BackgroundColor) {...
OutWord.Internal.ps1 GetWordStyle             613 $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xm...
OutWord.Internal.ps1 GetWordStyle             614 [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear')
OutWord.Internal.ps1 GetWordStyle             615 [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto')
OutWord.Internal.ps1 GetWordStyle             616 [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWord...
OutWord.Internal.ps1 GetWordStyle             616 ConvertToWordColor -Color $Style.BackgroundColor
OutWord.Internal.ps1 GetWordStyle             618 [ref] $null = $documentStyle.AppendChild((GetWordStyleRunPr -Style...
OutWord.Internal.ps1 GetWordStyle             618 GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument
OutWord.Internal.ps1 GetWordStyle             620 return $documentStyle
OutWord.Internal.ps1 GetWordTableStyle        636 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 GetWordTableStyle        637 $style = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
OutWord.Internal.ps1 GetWordTableStyle        638 [ref] $null = $style.SetAttribute('type', $xmlnsMain, 'table')
OutWord.Internal.ps1 GetWordTableStyle        639 [ref] $null = $style.SetAttribute('styleId', $xmlnsMain, $TableSty...
OutWord.Internal.ps1 GetWordTableStyle        640 $name = $style.AppendChild($XmlDocument.CreateElement('w', 'name',...
OutWord.Internal.ps1 GetWordTableStyle        641 [ref] $null = $name.SetAttribute('val', $xmlnsMain, $TableStyle.Id)
OutWord.Internal.ps1 GetWordTableStyle        642 $tblPr = $style.AppendChild($XmlDocument.CreateElement('w', 'tblPr...
OutWord.Internal.ps1 GetWordTableStyle        643 $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateEleme...
OutWord.Internal.ps1 GetWordTableStyle        644 [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlnsMain, 1)
OutWord.Internal.ps1 GetWordTableStyle        645 if ($tableStyle.BorderWidth -gt 0) {...
OutWord.Internal.ps1 GetWordTableStyle        646 $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', '...
OutWord.Internal.ps1 GetWordTableStyle        647 @('top','bottom','start','end','insideH','insideV')
OutWord.Internal.ps1 GetWordTableStyle        647 'top','bottom','start','end','insideH','insideV'
OutWord.Internal.ps1 GetWordTableStyle        648 $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $bord...
OutWord.Internal.ps1 GetWordTableStyle        649 [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips...
OutWord.Internal.ps1 GetWordTableStyle        649 ConvertMmToOctips $tableStyle.BorderWidth
OutWord.Internal.ps1 GetWordTableStyle        650 [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single')
OutWord.Internal.ps1 GetWordTableStyle        651 [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordC...
OutWord.Internal.ps1 GetWordTableStyle        651 ConvertToWordColor -Color $tableStyle.BorderColor
OutWord.Internal.ps1 GetWordTableStyle        654 [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Docu...
OutWord.Internal.ps1 GetWordTableStyle        654 GetWordTableStylePr -Style $Document.Styles[$TableStyle.HeaderStyl...
OutWord.Internal.ps1 GetWordTableStyle        655 [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Docu...
OutWord.Internal.ps1 GetWordTableStyle        655 GetWordTableStylePr -Style $Document.Styles[$TableStyle.RowStyle] ...
OutWord.Internal.ps1 GetWordTableStyle        656 [ref] $null = $style.AppendChild((GetWordTableStylePr -Style $Docu...
OutWord.Internal.ps1 GetWordTableStyle        656 GetWordTableStylePr -Style $Document.Styles[$TableStyle.AlternateR...
OutWord.Internal.ps1 GetWordTableStyle        657 return $style
OutWord.Internal.ps1 GetWordStyleParagraphPr  673 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 GetWordStyleParagraphPr  674 $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlnsMain)
OutWord.Internal.ps1 GetWordStyleParagraphPr  675 $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spaci...
OutWord.Internal.ps1 GetWordStyleParagraphPr  676 [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
OutWord.Internal.ps1 GetWordStyleParagraphPr  677 [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
OutWord.Internal.ps1 GetWordStyleParagraphPr  678 $keepNext = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keep...
OutWord.Internal.ps1 GetWordStyleParagraphPr  679 $keepLines = $pPr.AppendChild($XmlDocument.CreateElement('w', 'kee...
OutWord.Internal.ps1 GetWordStyleParagraphPr  680 $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmln...
OutWord.Internal.ps1 GetWordStyleParagraphPr  681 if ($Style.Align.ToLower() -eq 'justify') { [ref] $null = $jc.SetA...
OutWord.Internal.ps1 GetWordStyleParagraphPr  681 [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
OutWord.Internal.ps1 GetWordStyleParagraphPr  682 [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToL...
OutWord.Internal.ps1 GetWordStyleParagraphPr  683 return $pPr
OutWord.Internal.ps1 GetWordStyleRunPrColor   702 $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain)
OutWord.Internal.ps1 GetWordStyleRunPrColor   703 $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color',...
OutWord.Internal.ps1 GetWordStyleRunPrColor   704 [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWor...
OutWord.Internal.ps1 GetWordStyleRunPrColor   704 ConvertToWordColor -Color $Style.Color
OutWord.Internal.ps1 GetWordStyleRunPrColor   705 return $rPr
OutWord.Internal.ps1 GetWordStyleRunPr        721 $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain)
OutWord.Internal.ps1 GetWordStyleRunPr        722 $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts...
OutWord.Internal.ps1 GetWordStyleRunPr        723 [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Style.Fon...
OutWord.Internal.ps1 GetWordStyleRunPr        724 [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Style.Fon...
OutWord.Internal.ps1 GetWordStyleRunPr        725 if ($Style.Bold) {...
OutWord.Internal.ps1 GetWordStyleRunPr        726 [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b'...
OutWord.Internal.ps1 GetWordStyleRunPr        728 if ($Style.Underline) {...
OutWord.Internal.ps1 GetWordStyleRunPr        729 [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u'...
OutWord.Internal.ps1 GetWordStyleRunPr        731 if ($Style.Italic) {...
OutWord.Internal.ps1 GetWordStyleRunPr        732 [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i'...
OutWord.Internal.ps1 GetWordStyleRunPr        734 $color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color',...
OutWord.Internal.ps1 GetWordStyleRunPr        735 [ref] $null = $color.SetAttribute('val', $xmlnsMain, (ConvertToWor...
OutWord.Internal.ps1 GetWordStyleRunPr        735 ConvertToWordColor -Color $Style.Color
OutWord.Internal.ps1 GetWordStyleRunPr        736 $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmln...
OutWord.Internal.ps1 GetWordStyleRunPr        737 [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Style.Size * 2)
OutWord.Internal.ps1 GetWordStyleRunPr        738 return $rPr
OutWord.Internal.ps1 GetWordTableStylePr      779 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 GetWordTableStylePr      780 $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlns...
OutWord.Internal.ps1 GetWordTableStylePr      781 $tblPr = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', '...
OutWord.Internal.ps1 GetWordTableStylePr      782 $Type
OutWord.Internal.ps1 GetWordTableStylePr      783 $tblStylePrType = 'firstRow'
OutWord.Internal.ps1 GetWordTableStylePr      784 $tblStylePrType = 'band2Horz'
OutWord.Internal.ps1 GetWordTableStylePr      785 $tblStylePrType = 'band1Horz'
OutWord.Internal.ps1 GetWordTableStylePr      787 [ref] $null = $tblStylePr.SetAttribute('type', $xmlnsMain, $tblSty...
OutWord.Internal.ps1 GetWordTableStylePr      788 [ref] $null = $tblStylePr.AppendChild((GetWordStyleParagraphPr -St...
OutWord.Internal.ps1 GetWordTableStylePr      788 GetWordStyleParagraphPr -Style $Style -XmlDocument $XmlDocument
OutWord.Internal.ps1 GetWordTableStylePr      789 [ref] $null = $tblStylePr.AppendChild((GetWordStyleRunPr -Style $S...
OutWord.Internal.ps1 GetWordTableStylePr      789 GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument
OutWord.Internal.ps1 GetWordTableStylePr      790 [ref] $null = $tblStylePr.AppendChild((GetWordTableStyleCellPr -St...
OutWord.Internal.ps1 GetWordTableStylePr      790 GetWordTableStyleCellPr -Style $Style -XmlDocument $XmlDocument
OutWord.Internal.ps1 GetWordTableStylePr      791 return $tblStylePr
OutWord.Internal.ps1 GetWordSectionPr         812 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 GetWordSectionPr         813 $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlnsMain)
OutWord.Internal.ps1 GetWordSectionPr         814 $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz'...
OutWord.Internal.ps1 GetWordSectionPr         815 [ref] $null = $pgSz.SetAttribute('w', $xmlnsMain, (ConvertMmToTwip...
OutWord.Internal.ps1 GetWordSectionPr         815 ConvertMmToTwips -Millimeter $PageWidth
OutWord.Internal.ps1 GetWordSectionPr         816 [ref] $null = $pgSz.SetAttribute('h', $xmlnsMain, (ConvertMmToTwip...
OutWord.Internal.ps1 GetWordSectionPr         816 ConvertMmToTwips -Millimeter $PageHeight
OutWord.Internal.ps1 GetWordSectionPr         817 [ref] $null = $pgSz.SetAttribute('orient', $xmlnsMain, 'portrait')
OutWord.Internal.ps1 GetWordSectionPr         818 $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMa...
OutWord.Internal.ps1 GetWordSectionPr         819 [ref] $null = $pgMar.SetAttribute('top', $xmlnsMain, (ConvertMmToT...
OutWord.Internal.ps1 GetWordSectionPr         819 ConvertMmToTwips -Millimeter $PageMarginTop
OutWord.Internal.ps1 GetWordSectionPr         820 [ref] $null = $pgMar.SetAttribute('bottom', $xmlnsMain, (ConvertMm...
OutWord.Internal.ps1 GetWordSectionPr         820 ConvertMmToTwips -Millimeter $PageMarginBottom
OutWord.Internal.ps1 GetWordSectionPr         821 [ref] $null = $pgMar.SetAttribute('left', $xmlnsMain, (ConvertMmTo...
OutWord.Internal.ps1 GetWordSectionPr         821 ConvertMmToTwips -Millimeter $PageMarginLeft
OutWord.Internal.ps1 GetWordSectionPr         822 [ref] $null = $pgMar.SetAttribute('right', $xmlnsMain, (ConvertMmT...
OutWord.Internal.ps1 GetWordSectionPr         822 ConvertMmToTwips -Millimeter $PageMarginRight
OutWord.Internal.ps1 GetWordSectionPr         823 return $sectPr
OutWord.Internal.ps1 OutWordStylesDocument    842 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 OutWordStylesDocument    843 $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
OutWord.Internal.ps1 OutWordStylesDocument    844 [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDecla...
OutWord.Internal.ps1 OutWordStylesDocument    845 $documentStyles = $xmlDocument.AppendChild($xmlDocument.CreateElem...
OutWord.Internal.ps1 OutWordStylesDocument    848 $defaultStyle = $documentStyles.AppendChild($xmlDocument.CreateEle...
OutWord.Internal.ps1 OutWordStylesDocument    849 [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'para...
OutWord.Internal.ps1 OutWordStylesDocument    850 [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1')
OutWord.Internal.ps1 OutWordStylesDocument    851 [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'N...
OutWord.Internal.ps1 OutWordStylesDocument    852 $defaultStyleName = $defaultStyle.AppendChild($xmlDocument.CreateE...
OutWord.Internal.ps1 OutWordStylesDocument    853 [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'N...
OutWord.Internal.ps1 OutWordStylesDocument    854 [ref] $null = $defaultStyle.AppendChild($xmlDocument.CreateElement...
OutWord.Internal.ps1 OutWordStylesDocument    856 $Styles.Values
OutWord.Internal.ps1 OutWordStylesDocument    857 $documentParagraphStyle = GetWordStyle -Style $style -XmlDocument ...
OutWord.Internal.ps1 OutWordStylesDocument    858 [ref] $null = $documentStyles.AppendChild($documentParagraphStyle)
OutWord.Internal.ps1 OutWordStylesDocument    859 $documentCharacterStyle = GetWordStyle -Style $style -XmlDocument ...
OutWord.Internal.ps1 OutWordStylesDocument    860 [ref] $null = $documentStyles.AppendChild($documentCharacterStyle)
OutWord.Internal.ps1 OutWordStylesDocument    862 $TableStyles.Values
OutWord.Internal.ps1 OutWordStylesDocument    863 $documentTableStyle = GetWordTableStyle -TableStyle $tableStyle -X...
OutWord.Internal.ps1 OutWordStylesDocument    864 [ref] $null = $documentStyles.AppendChild($documentTableStyle)
OutWord.Internal.ps1 OutWordStylesDocument    866 return $xmlDocument
OutWord.Internal.ps1 OutWordSettingsDocument  882 $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2...
OutWord.Internal.ps1 OutWordSettingsDocument  893 $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument'
OutWord.Internal.ps1 OutWordSettingsDocument  894 [ref] $null = $settingsDocument.AppendChild($settingsDocument.Crea...
OutWord.Internal.ps1 OutWordSettingsDocument  895 $settings = $settingsDocument.AppendChild($settingsDocument.Create...
OutWord.Internal.ps1 OutWordSettingsDocument  897 $compat = $settings.AppendChild($settingsDocument.CreateElement('w...
OutWord.Internal.ps1 OutWordSettingsDocument  898 $compatSetting = $compat.AppendChild($settingsDocument.CreateEleme...
OutWord.Internal.ps1 OutWordSettingsDocument  899 [ref] $null = $compatSetting.SetAttribute('name', $xmlnsMain, 'com...
OutWord.Internal.ps1 OutWordSettingsDocument  900 [ref] $null = $compatSetting.SetAttribute('uri', $xmlnsMain, 'http...
OutWord.Internal.ps1 OutWordSettingsDocument  901 [ref] $null = $compatSetting.SetAttribute('val', $xmlnsMain, 15)
OutWord.Internal.ps1 OutWordSettingsDocument  902 if ($UpdateFields) {...
OutWord.Internal.ps1 OutWordSettingsDocument  903 $wupdateFields = $settings.AppendChild($settingsDocument.CreateEle...
OutWord.Internal.ps1 OutWordSettingsDocument  904 [ref] $null = $wupdateFields.SetAttribute('val', $xmlnsMain, 'true')
OutWord.Internal.ps1 OutWordSettingsDocument  906 return $settingsDocument
#>

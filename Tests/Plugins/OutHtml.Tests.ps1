$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'OutHtml\OutHtml' {
        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'warns when 7 nested sections are defined' {
            $testDocument = Document -Name 'IllegalNestedSections' -ScriptBlock {
                Section -Name 'Level1' {
                    Section -Name 'Level2' {
                        Section -Name 'Level3' {
                            Section -Name 'Level4' {
                                Section -Name 'Level5' {
                                    Section -Name 'Level6' {
                                        Section -Name 'Level7' { }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            { $testDocument | OutHtml -Path $path -WarningAction Stop 3>&1 } | Should Throw '6 heading'
        }

        It 'calls OutHtmlSection' {
            Mock -CommandName OutHtmlSection -Verifiable -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlSection -Exactly 1;
        }

        It 'calls OutHtmlParagraph' {
            Mock -CommandName OutHtmlParagraph -Verifiable -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlParagraph -Exactly 1;
        }

        It 'calls OutHtmlTable' {
            Mock -CommandName OutHtmlTable -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Get-Service | Select-Object -First 1 | Table 'TestTable' } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlTable -Exactly 1;
        }

        It 'calls OutHtmlLineBreak' {
            Mock -CommandName OutHtmlLineBreak -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { LineBreak; } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlLineBreak -Exactly 1;
        }

        It 'calls OutHtmlPageBreak' {
            Mock -CommandName OutHtmlPageBreak -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { PageBreak; } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlPageBreak -Exactly 1;
        }

        It 'calls OutHtmlTOC' {
            Mock -CommandName OutHtmlTOC -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { TOC -Name 'TestTOC'; } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlTOC -Exactly 1;
        }

        It 'calls OutHtmlBlankLine' {
            Mock -CommandName OutHtmlBlankLine -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlBlankLine -Exactly 1;
        }

        It 'calls OutHtmlBlankLine twice' {
            Mock -CommandName OutHtmlBlankLine -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; BlankLine; } | OutHtml -Path $path;
            Assert-MockCalled -CommandName OutHtmlBlankLine -Exactly 3; ## Mock calls are cumalative
        }

    }

    Describe 'OutHtml.Internal\GetHtmlStyle' {
        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        Context 'By named parameter.' {

            It 'creates single font default style.' {
                Style -Name Test -Font Helvetica;
                $fontFamily = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-family:*' ;
                ($fontFamily.Split(':').Trim())[1] | Should BeExactly "'Helvetica'";
                $fontSize = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*' ;
                ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92em';
                $fontWeight = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-weight:*' ;
                ($fontWeight.Split(':').Trim())[1] | Should BeExactly 'normal';
                $fontStyle = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-style:*' ;
                $fontStyle | Should BeNullOrEmpty;
                $textDecoration = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-decoration:*' ;
                $textDecoration | Should BeNullOrEmpty;
                $textAlign = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*' ;
                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'left';
                $color = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*' ;
                ($color.Split(':').Trim())[1] | Should BeExactly '#000000';
                $backgroundColor = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*' ;
                $backgroundColor | Should BeNullOrEmpty;
            }

            It 'uses invariant culture font size (#6)' {
                Style -Name Test -Font Helvetica;
                $currentCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture.Name;
                [System.Threading.Thread]::CurrentThread.CurrentCulture = 'da-DK';
                $fontSize = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*' ;
                [System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture;
                ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92em';
            }

            It 'creates multiple font default style.' {
                Style -Name Test -Font Helvetica,Arial,Sans-Serif;
                $fontFamily = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-family:*' ;
                ($fontFamily.Split(':').Trim())[1] | Should BeExactly "'Helvetica','Arial','Sans-Serif'";
            }

            It 'creates single 12pt font.' {
                Style -Name Test -Font Helvetica -Size 12;
                $fontSize = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*' ;
                ($fontSize.Split(':').Trim())[1] | Should BeExactly '1.00em';
            }

            It 'creates bold font style.' {
                Style -Name Test -Font Helvetica -Bold;
                $fontWeight = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-weight:*' ;
                ($fontWeight.Split(':').Trim())[1] | Should BeExactly 'bold';
            }

            It 'creates center aligned font style.' {
                Style -Name Test -Font Helvetica -Align Center;
                $textAlign = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*' ;
                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'center';
            }

            It 'creates right aligned font style.' {
                Style -Name Test -Font Helvetica -Align Right;
                $textAlign = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*' ;
                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'right';
            }

            It 'creates justified font style.' {
                Style -Name Test -Font Helvetica -Align Justify;
                $textAlign = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*' ;
                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'justify';
            }

            It 'creates underline font style.' {
                Style -Name Test -Font Helvetica -Underline;
                $textDecoration = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-decoration:*' ;
                ($textDecoration.Split(':').Trim())[1] | Should BeExactly 'underline';
            }

            It 'creates italic font style.' {
                Style -Name Test -Font Helvetica -Italic;
                $fontStyle = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-style:*' ;
                ($fontStyle.Split(':').Trim())[1] | Should BeExactly 'italic';
            }

            It 'creates colored font style.' {
                Style -Name Test -Font Helvetica -Color ABC;
                $color = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*' ;
                ($color.Split(':').Trim())[1] | Should BeExactly '#abc';
            }

            It 'creates colored font style with #.' {
                Style -Name Test -Font Helvetica -Color '#ABC';
                $color = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*' ;
                ($color.Split(':').Trim())[1] | Should BeExactly '#abc';
            }

            It 'creates background colored font style.' {
                Style -Name Test -Font Helvetica -BackgroundColor '#DEF';
                $backgroundColor = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*' ;
                ($backgroundColor.Split(':').Trim())[1] | Should BeExactly '#def';

            }

            It 'creates background colored font without #.' {
                Style -Name Test -Font Helvetica -BackgroundColor 'DEF';
                $backgroundColor = ((GetHtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*' ;
                ($backgroundColor.Split(':').Trim())[1] | Should BeExactly '#def';
            }

        } #end context By Named Parameter

    } #end describe GetHtmlStyle

    Describe 'OutHtml.Internal\GetHtmlTableStyle' {
        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        Context 'By Named Parameter.' {

            It 'creates default table style.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle;

                $padding = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'padding:*' ;
                ($padding.Split(':').Trim())[1] | Should BeExactly '0.08em 0.33em 0em 0.33em';
                #$borderColor = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*' ;
                #($borderColor.Split(':').Trim())[1] | Should BeExactly '#000';
                #$borderWidth = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-width:*' ;
                #($borderWidth.Split(':').Trim())[1] | Should BeExactly '0em';
                $borderStyle = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*' ;
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'none';
                $borderCollapse = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-collapse:*' ;
                ($borderCollapse.Split(':').Trim())[1] | Should BeExactly 'collapse';
            }

            It 'creates custom table padding style of 5pt, 10pt, 5pt and 10pt.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -PaddingTop 5 -PaddingRight 10 -PaddingBottom 5 -PaddingLeft 10;

                $padding = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'padding:*' ;
                ($padding.Split(':').Trim())[1] | Should BeExactly '0.42em 0.83em 0.42em 0.83em';
            }

            It 'creates custom table border color style when -BorderWidth is specified.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -BorderColor CcC -BorderWidth 1;

                $borderColor = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*' ;
                ($borderColor.Split(':').Trim())[1] | Should BeExactly '#ccc';
                $borderStyle = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*' ;
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid';
            }

             It 'creates custom table border width style of 3pt.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -BorderWidth 3;

                $borderWidth = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-width:*' ;
                ($borderWidth.Split(':').Trim())[1] | Should BeExactly '0.25em';
                $borderStyle = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*' ;
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid';
            }

            It 'creates custom table border with no color style when no -BorderWidth specified.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -BorderColor '#aAaAaA';

                $borderStyle = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*' ;
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'none';
            }

            It 'creates custom table border color style.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -BorderColor '#aAaAaA' -BorderWidth 2;

                $borderColor = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*' ;
                ($borderColor.Split(':').Trim())[1] | Should BeExactly '#aaaaaa';
                $borderStyle = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*' ;
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid';
            }

            It 'centers table.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -Align Center;
                $marginLeft = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-left:*' ;
                $marginRight = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-right:*' ;
                ($marginLeft.Split(':').Trim())[1] | Should BeExactly 'auto';
                ($marginRight.Split(':').Trim())[1] | Should BeExactly 'auto';
            }

            It 'aligns table to the right.' {
                Style -Name Default -Font Helvetica -Default;
                TableStyle -Name TestTableStyle -Align Right;
                $marginLeft = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-left:*' ;
                $marginRight = ((GetHtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-right:*' ;
                ($marginLeft.Split(':').Trim())[1] | Should BeExactly 'auto';
                ($marginRight.Split(':').Trim())[1] | Should BeExactly '0';
            }

        } #end context By Named Parameter

    } #end describe GetHtmlTableStyle

    Describe 'OutHtml.Internal\OutHtmlBlankLine' {
        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        It 'creates a single <br /> html tag.' {
            BlankLine | OutHtmlBlankLine | Should BeExactly '<br />';
        }

        It 'creates two <br /> html tags.' {
            BlankLine -Count 2 | OutHtmlBlankLine | Should BeExactly '<br /><br />';
        }
    }

    Describe 'OutHtml.Internal\OutHtmlLineBreak' {

        It 'creates a <hr /> html tag.' {
            OutHtmlLineBreak | Should BeExactly '<hr />';
        }
    }

    Describe 'OutHtml.Internal\OutHtmlPageBreak' {
        ## Scaffold new document to initialise options/styles
        $Document = Document -Name 'Test' -ScriptBlock { };
        $text = OutHtmlPageBreak;

        It 'closes previous </page> and </div> tags.' {
            $text.StartsWith('</div></page>') | Should Be $true;
        }

        It 'creates new <page>.' {
            $text -match '<page>' | Should Be $true;
        }

        It 'sets page class to default style' {
            $divStyleMatch = '<div class="{0}"' -f $Document.DefaultStyle;
            $text -match $divStyleMatch | Should Be $true;
        }

        It 'includes page margins.' {
            $text -match 'padding-top:[\s\S]+em' | Should Be $true;
            $text -match 'padding-right:[\s\S]+em' | Should Be $true;
            $text -match 'padding-bottom:[\s\S]+em' | Should Be $true;
            $text -match 'padding-left:[\s\S]+em' | Should Be $true;
        }

    }

    Describe 'OutHtml.Internal\OutHtmlParagraph' {
        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        Context 'By Named Parameter.' {

            It 'creates paragraph with no style and new line.' {
                Paragraph 'Test paragraph.' | OutHtmlParagraph | Should BeExactly "<div>Test paragraph.</div>";
            }

            It 'creates paragraph with custom name/id' {
                Paragraph -Name 'Test' -Text 'Test paragraph.' -NoNewLine | OutHtmlParagraph | Should BeExactly "<div>Test paragraph.</div>";
            }

            It 'creates paragraph with named -Style parameter.' {
                Paragraph 'Test paragraph.' -Style Named | OutHtmlParagraph | Should BeExactly "<div class=`"Named`">Test paragraph.</div>";
            }

            It 'encodes HTML paragraph content' {
                $expected = '<div>Embedded &lt;br /&gt;</div>';
                $result = Paragraph 'Embedded <br />' | OutHtmlParagraph;
                $result | Should BeExactly $expected;
            }

            It 'creates paragraph with embedded new line' {
                $expected = '<div>Embedded<br />New Line</div>';
                $result = Paragraph "Embedded`r`nNew Line" | OutHtmlParagraph;
                $result | Should BeExactly $expected;
            }

        } #end context By Named Parameter

    } #end describe OutHtmlParagraph

    Describe 'OutHtml.Internal\OutHtmlSection' {
        $Document = Document -Name 'TestDocument' -ScriptBlock { }
        $pscriboDocument = $Document;

        It 'calls OutHtmlParagraph' {
            Mock -CommandName OutHtmlParagraph -MockWith { };
            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlParagraph -Exactly 1;
        }

        It 'calls OutHtmlParagraph twice' {
            Mock -CommandName OutHtmlParagraph -MockWith { };
            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph'; Paragraph 'TestParagraph'; } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlParagraph -Exactly 3;
        }

        It 'calls OutHtmlTable' {
            Mock -CommandName OutHtmlTable -MockWith { };
            Section -Name TestSection -ScriptBlock { Get-Service | Select-Object -First 3 | Table TestTable } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlTable -Exactly 1;
        }

        It 'calls OutHtmlPageBreak' {
            Mock -CommandName OutHtmlPageBreak -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { PageBreak } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlPageBreak -Exactly 1;
        }

        It 'calls OutHtmlLineBreak' {
            Mock -CommandName OutHtmlLineBreak -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { LineBreak } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlLineBreak -Exactly 1;
        }

        It 'calls OutHtmlBlankLine' {
            Mock -CommandName OutHtmlBlankLine -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { BlankLine } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlBlankLine -Exactly 1;
        }

        It 'warns on call OutHtmlTOC' {
            Mock -CommandName OutHtmlTOC -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { TOC 'TestTOC' } | OutHtmlSection -WarningAction SilentlyContinue;
            Assert-MockCalled OutHtmlTOC -Exactly 0;
        }

        It 'encodes HTML section name' {
            $sectionName = 'Test & Section';
            $expected = '<h1 class="Normal">{0}</h1>' -f $sectionName.Replace('&','&amp;');

            $result = Section -Name $sectionName -ScriptBlock { BlankLine } | OutHtmlSection;

            $result -match $expected | Should Be $true;
        }

        It 'calls nested OutXmlSection' {
            ## Note this must be called last in the Describe script block as the OutXmlSection gets mocked!
            Mock -CommandName OutHtmlSection -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlSection -Exactly 1;
        }

    }

    Describe 'OutHtml.Internal\OutHtmlStyle' {

        It 'creates <style> tag.' {
            $Document = Document -Name 'Test' -ScriptBlock { };
            $text = OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles;

            $text -match '<style type="text/css">' | Should Be $true;
            $text -match '</style>' | Should Be $true;
        }

        It 'creates page layout style by default' {
            $Document = Document -Name 'Test' -ScriptBlock { };
            $text = OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles;

            $text -match 'html {' | Should Be $true;
            $text -match 'page {' | Should Be $true;
            $text -match '@media print {' | Should Be $true;
        }

        It "suppresses page layout style when 'Options.NoPageLayoutSyle' specified" {
            $Document = Document -Name 'Test' -ScriptBlock { };
            $text = OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles -NoPageLayoutStyle;

            $text -match 'html {' | Should Be $false;
            $text -match 'page {' | Should Be $false;
            $text -match '@media print {' | Should Be $false;
        }
    }

    Describe 'OutHtml.Internal\OutHtmlParagraphStyle' {

        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        It 'uses invariant culture paragraph size (#6)' {
            $currentCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture.Name;
            [System.Threading.Thread]::CurrentThread.CurrentCulture = 'da-DK';

            $result = (Paragraph 'Test paragraph.' -Size 11 | OutHtmlParagraph) -match '(?<=style=").+(?=">)';
            $fontSize = ($Matches[0]).Trim(';');

            ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92em';
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture;
        }

    } #end describe OutHtmlParagraphStyle

    Describe 'OutHtml.Internal\OutHtmlTable' {

        Context 'Table.' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { };
                $services = Get-Service | Select-Object -First 3;
                $table = $services | Table -Name 'Test Table' | OutHtmlTable;
                [Xml] $html = $table.Replace('&','&amp;');
            }

            It 'creates default table class of tabledefault.' {
                $html.Div.Table.Class | Should BeExactly 'tabledefault';
            }

            It 'creates table headings row.' {
                $html.Div.Table.Thead | Should Not BeNullOrEmpty;
            }

            It 'creates column for each object property.' {
                $html.Div.Table.Thead.Tr.Th.Count | Should Be ($services | Get-Member -MemberType Properties).Count;
            }

            It 'creates a row for each object.' {
                $html.Div.Table.Tbody.Tr.Count | Should Be $services.Count;
            }

        }

        Context 'List.' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { };
                $services = Get-Service | Select -First 1;
                $table = $services | Table -Name 'Test Table' -List | OutHtmlTable;
                [Xml] $html = $table.Replace('&','&amp;');
            }

            #Write-Host $html.OuterXml -ForegroundColor Yellow

            It 'creates no table heading row.' {
                ## Fix Set-StrictMode
                $html.Div.Table.PSObject.Properties['Thead'] | Should BeNullOrEmpty;
            }

            It 'creates default table class of tabledefault-list.' {
                $html.Div.Table.Class | Should BeExactly 'tabledefault-list';
            }

            It 'creates a two column table.' {
                $html.Div.Table.Tbody.Tr[0].Td.Count | Should Be 2;
            }

            It 'creates a row for each object property.' {
                $html.Div.Table.Tbody.Tr.Count | Should Be ($services | Get-Member -MemberType Properties).Count;
            }

        } #end context List

        Context 'New Lines' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { };
            }

            It 'creates a tabular table cell with an embedded new line' {

                $licenses = "Standard`r`nProfessional`r`nEnterprise"
                $expected = '<td>Standard<br />Professional<br />Enterprise</td>';
                $newLineTable = [PSCustomObject] @{ 'Licenses' = $licenses; }

                $table = $newLineTable | Table -Name 'Test Table' | OutHtmlTable;

                [Xml] $html = $table.Replace('&','&amp;');
                $html.OuterXml | Should Match $expected;
            }

            It 'creates a list table cell with an embedded new line' {

                $licenses = "Standard`r`nProfessional`r`nEnterprise"
                $expected = '<td>Standard<br />Professional<br />Enterprise</td>';
                $newLineTable = [PSCustomObject] @{ 'Licenses' = $licenses; }

                $table = $newLineTable | Table -Name 'Test Table' | OutHtmlTable;

                [Xml] $html = $table.Replace('&','&amp;');
                $html.OuterXml | Should Match $expected;

            }

        }

    } #end describe OutHtmlTable

} #end inmodulescope

<#
Code coverage report:
Covered 78.87% of 265 analyzed commands in 1 file.

Missed commands:

File                 Function              Line Command
----                 --------              ---- -------
OutHtml.Internal.ps1 GetHtmlStyle            25 [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Colo...
OutHtml.Internal.ps1 GetHtmlStyle            28 [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', ...
OutHtml.Internal.ps1 GetHtmlTableStyle       58 [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: {0};',...
OutHtml.Internal.ps1 GetHtmlTableDiv         89 [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}...
OutHtml.Internal.ps1 GetHtmlTableDiv         89 ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)
OutHtml.Internal.ps1 GetHtmlTableDiv         89 12.7 * $Table.Tabs
OutHtml.Internal.ps1 GetHtmlTableDiv        105 $styleElements += 'table-layout: fixed;'
OutHtml.Internal.ps1 GetHtmlTableDiv        106 $styleElements += 'word-break: break-word;'
OutHtml.Internal.ps1 GetHtmlTableDiv        112 [ref] $null = $divBuilder.Append('>')
OutHtml.Internal.ps1 GetHtmlTableColGroup   131 [ref] $null = $colGroupBuilder.Append('<colgroup>')
OutHtml.Internal.ps1 GetHtmlTableColGroup   132 $Table.ColumnWidths
OutHtml.Internal.ps1 GetHtmlTableColGroup   133 if ($null -eq $columnWidth) {...
OutHtml.Internal.ps1 GetHtmlTableColGroup   134 [ref] $null = $colGroupBuilder.Append('<col />')
OutHtml.Internal.ps1 GetHtmlTableColGroup   137 [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{...
OutHtml.Internal.ps1 GetHtmlTableColGroup   140 [ref] $null = $colGroupBuilder.AppendLine('</colgroup>')
OutHtml.Internal.ps1 OutHtmlTOC             157 $tocBuilder = New-Object -TypeName 'System.Text.StringBuilder'
OutHtml.Internal.ps1 OutHtmlTOC             158 [ref] $null = $tocBuilder.AppendFormat('<h1 class="TOC">{0}</h1>', $...
OutHtml.Internal.ps1 OutHtmlTOC             160 [ref] $null = $tocBuilder.AppendLine('<table>')
OutHtml.Internal.ps1 OutHtmlTOC             161 $Document.TOC
OutHtml.Internal.ps1 OutHtmlTOC             162 $sectionNumberIndent = '&nbsp;&nbsp;&nbsp;' * $tocEntry.Level
OutHtml.Internal.ps1 OutHtmlTOC             163 if ($Document.Options['EnableSectionNumbering']) {...
OutHtml.Internal.ps1 OutHtmlTOC             164 [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a hr...
OutHtml.Internal.ps1 OutHtmlTOC             167 [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}<a href="#{1}" st...
OutHtml.Internal.ps1 OutHtmlTOC             170 [ref] $null = $tocBuilder.AppendLine('</table>')
OutHtml.Internal.ps1 OutHtmlTOC             171 return $tocBuilder.ToString()
OutHtml.Internal.ps1 OutHtmlStyle           215 $Document.Options['PageWidth']
OutHtml.Internal.ps1 OutHtmlSection         261 [string] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectio...
OutHtml.Internal.ps1 OutHtmlSection         266 WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning
OutHtml.Internal.ps1 OutHtmlSection         267 $headerLevel = 5
OutHtml.Internal.ps1 OutHtmlSection         270 $className = $Section.Style
OutHtml.Internal.ps1 OutHtmlSection         273 $sectionId = '{0}[..]' -f $s.Id.Substring(0,36)
OutHtml.Internal.ps1 OutHtmlSection         276 $currentIndentationLevel = $s.Level +1
OutHtml.Internal.ps1 OutHtmlSection         279 [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s))
OutHtml.Internal.ps1 OutHtmlSection         279 OutHtmlSection -Section $s
OutHtml.Internal.ps1 GetHtmlParagraphStyle  306 $tabEm = ConvertMmToEm -Millimeter (12.7 * $Paragraph.Tabs)
OutHtml.Internal.ps1 GetHtmlParagraphStyle  306 12.7 * $Paragraph.Tabs
OutHtml.Internal.ps1 GetHtmlParagraphStyle  307 [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}...
OutHtml.Internal.ps1 GetHtmlParagraphStyle  309 [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0...
OutHtml.Internal.ps1 GetHtmlParagraphStyle  315 [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;')
OutHtml.Internal.ps1 GetHtmlParagraphStyle  316 [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;')
OutHtml.Internal.ps1 GetHtmlParagraphStyle  317 [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: under...
OutHtml.Internal.ps1 GetHtmlParagraphStyle  319 [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $P...
OutHtml.Internal.ps1 GetHtmlParagraphStyle  322 [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $...
OutHtml.Internal.ps1 GetHtmlTableList       381 $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$pro...
OutHtml.Internal.ps1 GetHtmlTableList       381 GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle]
OutHtml.Internal.ps1 GetHtmlTableList       382 if ([string]::IsNullOrEmpty($Row.$propertyName)) {...
OutHtml.Internal.ps1 GetHtmlTableList       383 [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;...
OutHtml.Internal.ps1 GetHtmlTableList       386 [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</t...
OutHtml.Internal.ps1 GetHtmlTableList       386 $propertyName
OutHtml.Internal.ps1 GetHtmlTable           434 $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$pro...
OutHtml.Internal.ps1 GetHtmlTable           434 GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]
OutHtml.Internal.ps1 GetHtmlTable           435 [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1...
OutHtml.Internal.ps1 GetHtmlTable           439 $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style])...
OutHtml.Internal.ps1 GetHtmlTable           439 GetHtmlStyle -Style $Document.Styles[$row.__Style]
OutHtml.Internal.ps1 GetHtmlTable           440 [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1...
OutHtml.Internal.ps1 OutHtmlTable           479 [ref] $null = $tableBuilder.AppendLine('<p />')
#>

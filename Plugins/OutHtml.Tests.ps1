$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$moduleRoot = Split-Path -Path $here -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'OutHtml' {
        $path = (Get-PSDrive -Name TestDrive).Root;

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

    Describe 'GetHtmlStyle' {
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

    Describe 'GetHtmlTableStyle' {
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

    Describe 'OutHtmlBlankLine' {
        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        It 'creates a single <br /> html tag.' {
            BlankLine | OutHtmlBlankLine | Should BeExactly '<br />';
        }

        It 'creates two <br /> html tags.' {
            BlankLine -Count 2 | OutHtmlBlankLine | Should BeExactly '<br /><br />';
        }
    }

    Describe 'OutHtmlLineBreak' {

        It 'creates a <hr /> html tag.' {
            OutHtmlLineBreak | Should BeExactly '<hr />';
        }
    }

    Describe 'OutHtmlPageBreak' {
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

    Describe 'OutHtmlParagraph' {
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

        } #end context By Named Parameter

    } #end describe OutHtmlParagraph

    Describe 'OutHtmlSection' {
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

        It 'calls nested OutXmlSection' {
            ## Note this must be called last in the Describe script block as the OutXmlSection gets mocked!
            Mock -CommandName OutHtmlSection -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | OutHtmlSection;
            Assert-MockCalled -CommandName OutHtmlSection -Exactly 1;
        }
    
    }

    Describe 'OutHtmlStyle' {
        $Document = Document -Name 'Test' -ScriptBlock { };
        $text = OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles;
        ## Predominately calls GetHtmlStyle and GetHtmlTableStyle

        It 'creates <style> tag.' {
            $text -match '<style type="text/css">' | Should Be $true;
            $text -match '</style>' | Should Be $true;
        }
    }

    Describe 'OutHtmlTable' {

        Context 'Table.' {
            ## Scaffold new document to initialise options/styles
            $pscriboDocument = Document -Name 'Test' -ScriptBlock { };
            $services = Get-Service | Select -First 3;
            $table = $services | Table -Name 'Test Table' | OutHtmlTable;
            [Xml] $html = $table.Replace('&','&amp;');
            
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
            ## Scaffold new document to initialise options/styles
            $pscriboDocument = Document -Name 'Test' -ScriptBlock { };
            $services = Get-Service | Select -First 1;
            $table = $services | Table -Name 'Test Table' -List | OutHtmlTable;
            [Xml] $html = $table.Replace('&','&amp;');

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

    } #end describe OutHtmlTable

} #end inmodulescope

<#
Code coverage report:
Covered 81.08% of 259 analyzed commands in 1 file.

Missed commands:

File                 Function              Line Command                                                                                                                                                   
----                 --------              ---- -------                                                                                                                                                   
OutHtml.Internal.ps1 GetHtmlStyle            24 [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Color.ToLower())                                                                          
OutHtml.Internal.ps1 GetHtmlStyle            27 [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', $Style.BackgroundColor.ToLower())                                                     
OutHtml.Internal.ps1 GetHtmlTableDiv         74 [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}em;">' -f (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)))                           
OutHtml.Internal.ps1 GetHtmlTableDiv         74 ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)                                                                                                            
OutHtml.Internal.ps1 GetHtmlTableDiv         74 12.7 * $Table.Tabs                                                                                                                                        
OutHtml.Internal.ps1 GetHtmlTableDiv         90 $styleElements += 'table-layout: fixed;'                                                                                                                  
OutHtml.Internal.ps1 GetHtmlTableDiv         91 $styleElements += 'word-wrap: break-word;'                                                                                                                
OutHtml.Internal.ps1 GetHtmlTableDiv         97 [ref] $null = $divBuilder.Append('>')                                                                                                                     
OutHtml.Internal.ps1 GetHtmlTableColGroup   116 [ref] $null = $colGroupBuilder.Append('<colgroup>')                                                                                                       
OutHtml.Internal.ps1 GetHtmlTableColGroup   117 $Table.ColumnWidths                                                                                                                                       
OutHtml.Internal.ps1 GetHtmlTableColGroup   118 if ($null -eq $columnWidth) {...                                                                                                                          
OutHtml.Internal.ps1 GetHtmlTableColGroup   119 [ref] $null = $colGroupBuilder.Append('<col />')                                                                                                          
OutHtml.Internal.ps1 GetHtmlTableColGroup   122 [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{0}%; min-width:{0}%; width:{0}%" />', $columnWidth)                                   
OutHtml.Internal.ps1 GetHtmlTableColGroup   125 [ref] $null = $colGroupBuilder.AppendLine('</colgroup>')                                                                                                  
OutHtml.Internal.ps1 OutHtmlTOC             148 [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a href="#{2}" style="text-decoration: none;">{3}</a></td></tr>', $tocEntry.Number, $sec...
OutHtml.Internal.ps1 OutHtmlStyle           196 $Document.Options['PageWidth']                                                                                                                            
OutHtml.Internal.ps1 OutHtmlSection         240 [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name                                                                                       
OutHtml.Internal.ps1 OutHtmlSection         245 WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning                                                                                            
OutHtml.Internal.ps1 OutHtmlSection         246 $headerLevel = 5                                                                                                                                          
OutHtml.Internal.ps1 OutHtmlSection         249 $className = $Section.Style                                                                                                                               
OutHtml.Internal.ps1 OutHtmlSection         252 $sectionId = '{0}[..]' -f $s.Id.Substring(0,36)                                                                                                           
OutHtml.Internal.ps1 OutHtmlSection         256 [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s))                                                                                        
OutHtml.Internal.ps1 OutHtmlSection         256 OutHtmlSection -Section $s                                                                                                                                
OutHtml.Internal.ps1 GetHtmlParagraphStyle  283 $tabEm = ConvertMmToEm -Millimeter (12.7 * $Paragraph.Tabs)                                                                                               
OutHtml.Internal.ps1 GetHtmlParagraphStyle  283 12.7 * $Paragraph.Tabs                                                                                                                                    
OutHtml.Internal.ps1 GetHtmlParagraphStyle  284 [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}em;', $tabEm)                                                                         
OutHtml.Internal.ps1 GetHtmlParagraphStyle  286 [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0}';", $Paragraph.Font -Join "','")                                                    
OutHtml.Internal.ps1 GetHtmlParagraphStyle  287 [ref] $null = $paragraphStyleBuilder.AppendFormat(' font-size: {0:0.00}em;', $Paragraph.Size / 12)                                                        
OutHtml.Internal.ps1 GetHtmlParagraphStyle  288 [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;')                                                                                        
OutHtml.Internal.ps1 GetHtmlParagraphStyle  289 [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;')                                                                                       
OutHtml.Internal.ps1 GetHtmlParagraphStyle  290 [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: underline;')                                                                               
OutHtml.Internal.ps1 GetHtmlParagraphStyle  292 [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $Paragraph.Color.ToLower())                                                             
OutHtml.Internal.ps1 GetHtmlParagraphStyle  295 [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $Paragraph.Color.ToLower())                                                            
OutHtml.Internal.ps1 OutHtmlParagraph       325 [ref] $null = $paragraphBuilder.AppendFormat('<div style="{1}">{2}</div>', $Paragraph.Style, $customStyle, $text)                                         
OutHtml.Internal.ps1 GetHtmlTableList       352 $propertyDisplayName = $Table.Headers[$i]                                                                                                                 
OutHtml.Internal.ps1 GetHtmlTableList       358 $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle])                                                                          
OutHtml.Internal.ps1 GetHtmlTableList       358 GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle]                                                                                                 
OutHtml.Internal.ps1 GetHtmlTableList       359 if ([string]::IsNullOrEmpty($Row.$propertyName)) {...                                                                                                     
OutHtml.Internal.ps1 GetHtmlTableList       360 [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;</td></tr>', $propertyStyleHtml)                                                      
OutHtml.Internal.ps1 GetHtmlTableList       363 [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</td></tr>', $propertyStyleHtml, $Row.($propertyName))                                   
OutHtml.Internal.ps1 GetHtmlTableList       363 $propertyName                                                                                                                                             
OutHtml.Internal.ps1 GetHtmlTable           402 [ref] $null = $standardTableBuilder.AppendFormat('<th>{0}</th>', $Table.Headers[$i])                                                                      
OutHtml.Internal.ps1 GetHtmlTable           415 $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]).Trim()                                                                   
OutHtml.Internal.ps1 GetHtmlTable           415 GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]                                                                                                 
OutHtml.Internal.ps1 GetHtmlTable           416 [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $row.$propertyName)                                      
OutHtml.Internal.ps1 GetHtmlTable           420 $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style]).Trim()                                                                               
OutHtml.Internal.ps1 GetHtmlTable           420 GetHtmlStyle -Style $Document.Styles[$row.__Style]                                                                                                        
OutHtml.Internal.ps1 GetHtmlTable           421 [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $rowStyleHtml, $row.$propertyName)                                           
OutHtml.Internal.ps1 OutHtmlTable           460 [ref] $null = $tableBuilder.AppendLine('<p />')   
#>
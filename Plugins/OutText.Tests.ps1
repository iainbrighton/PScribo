$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$moduleRoot = Split-Path -Path $here -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'OutText' {
        $path = (Get-PSDrive -Name TestDrive).Root;

        It 'calls OutTextSection' {
            Mock -CommandName OutTextSection -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Section -Name 'TestSection' -ScriptBlock { } } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextSection -Exactly 1;
        }

        It 'calls OutTextParagraph' {
            Mock -CommandName OutTextParagraph -Verifiable -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Paragraph 'TestParagraph' } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextParagraph -Exactly 1;
        }

        It 'calls OutTextLineBreak' {
            Mock -CommandName OutTextLineBreak -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { LineBreak; } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextLineBreak -Exactly 1;
        }

        It 'calls OutTextPageBreak' {
            Mock -CommandName OutTextPageBreak -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { PageBreak; } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextPageBreak -Exactly 1;
        }

         It 'calls OutTextTable' {
            Mock -CommandName OutTextTable -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { Get-Service | Select-Object -First 1 | Table 'TestTable' } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextTable -Exactly 1;
        }

        It 'calls OutTextTOC' {
            Mock -CommandName OutTextTOC -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { TOC -Name 'TestTOC'; } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextTOC -Exactly 1;
        }

        It 'calls OutTextBlankLine' {
            Mock -CommandName OutTextBlankLine -MockWith { };
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextBlankLine -Exactly 1;
        }

        It 'calls OutTextBlankLine twice' {
            Document -Name 'TestDocument' -ScriptBlock { BlankLine; BlankLine; } | OutText -Path $path;
            Assert-MockCalled -CommandName OutTextBlankLine -Exactly 3; ## Mock calls are cumalative
        }

    }

    Describe 'OutTextBlankLine' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        It 'Defaults to a single blank line.' {
            $l = BlankLine | OutTextBlankLine;
            $l | Should Be "`r`n";
        }

        It 'Creates 3 blank lines.' {
            $l = BlankLine -Count 3 | OutTextBlankLine;
            $l | Should Be "`r`n`r`n`r`n";
        }

    }

    Describe 'OutTextLineBreak' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
    
        It 'Defaults to 120 and includes new line.' {
            $l = OutTextLineBreak;
            $l.Length | Should Be 122;
        }

        It 'Truncates to 40 and includes new line.' {
            $Options = New-PScriboTextOptions -TextWidth 40 -SeparatorWidth 40;
            $l = OutTextLineBreak;
            $l.Length | Should Be 42;
        }

        It 'Wraps lines and includes new line' {
            $Options = New-PScriboTextOptions -TextWidth 40 -SeparatorWidth 80;
            $l = OutTextLineBreak
            $l.Length | Should Be 84;
        }

    } #end describe OutTextLineBreak

    Describe 'OutTextPageBreak' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
    
        It 'Defaults to 120 and includes new line.' {
            #$Options = New-PScriboTextOptions;
            $l = OutTextPageBreak;
            $l.Length | Should Be 124;
        }

        It 'Truncates to 40 and includes new line.' {
            $Options = New-PScriboTextOptions -TextWidth 40 -SeparatorWidth 40;
            $l = OutTextPageBreak;
            $l.Length | Should Be 44;
        }

        It 'Wraps lines and includes new line.' {
            $Options = New-PScriboTextOptions -TextWidth 40 -SeparatorWidth 80;
            $l = OutTextPageBreak
            $l.Length | Should Be 86;
        }

    } #end describe OutTextLineBreak

    Describe 'OutTextParagraph' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        Context 'By pipeline.' {
     
            It 'Paragraph with new line.' {
                $testParagraph = 'Test paragraph.';
                $p = Paragraph $testParagraph | OutTextParagraph;
                $p | Should BeExactly "Test paragraph.`r`n";
            }

            It 'Paragraph with no new line.' {
                $testParagraph = 'Test paragraph.';
                $p = Paragraph $testParagraph -NoNewLine | OutTextParagraph;
                $p | Should BeExactly $testParagraph;
            }
        
            It 'Paragraph wraps at 10 characters with new line.' {
                $testParagraph = 'Test paragraph.';
                $Options = New-PScriboTextOptions -TextWidth 10;
                $p = Paragraph $testParagraph | OutTextParagraph;
                $p | Should BeExactly "Test parag`r`nraph.`r`n";
            }

             It 'Paragraph wraps at 10 characters with no new line.' {
                $testParagraph = 'Test paragraph.';
                $Options = New-PScriboTextOptions -TextWidth 10;
                $p = Paragraph $testParagraph -NoNewLine | OutTextParagraph;
                $p | Should BeExactly "Test parag`r`nraph.";
            }

        } #end context by pipeline

        Context 'By named -Paragraph parameter.' {
        
            It 'By named -Paragraph parameter with new line.' {
                $testParagraph = 'Test paragraph.';
                $p = OutTextParagraph -Paragraph (Paragraph $testParagraph);
                $p | Should BeExactly "$testParagraph`r`n";
            }

            It 'By named -Paragraph parameter with no new line.' {
                $testParagraph = 'Test paragraph.';
                $p = OutTextParagraph -Paragraph (Paragraph $testParagraph -NoNewLine);
                $p | Should BeExactly $testParagraph;
            }
        } #end context -paragraph

    } #end describe outtextparagraph

    Describe 'OutTextSection' {
        $Document = Document -Name 'TestDocument' -ScriptBlock { };
        $pscriboDocument = $Document;

        It 'calls OutTextParagraph' {
            Mock -CommandName OutTextParagraph -MockWith { };
            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph' } | OutTextSection;
            Assert-MockCalled -CommandName OutTextParagraph -Exactly 1;
        }

        It 'calls OutTextParagraph twice' {
            Mock -CommandName OutTextParagraph -MockWith { };
            Section -Name TestSection -ScriptBlock { Paragraph 'TestParagraph'; Paragraph 'TestParagraph'; } | OutTextSection;
            Assert-MockCalled -CommandName OutTextParagraph -Exactly 3;
        }

        It 'calls OutTextTable' {
            Mock -CommandName OutTextTable -MockWith { };
            Section -Name TestSection -ScriptBlock { Get-Service | Select-Object -First 3 | Table TestTable } | OutTextSection;
            Assert-MockCalled -CommandName OutTextTable -Exactly 1;
        }

        It 'calls OutTextPageBreak' {
            Mock -CommandName OutTextPageBreak -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { PageBreak } | OutTextSection;
            Assert-MockCalled -CommandName OutTextPageBreak -Exactly 1;
        }

         It 'calls OutTextLineBreak' {
            Mock -CommandName OutTextLineBreak -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { LineBreak } | OutTextSection;
            Assert-MockCalled -CommandName OutTextLineBreak -Exactly 1;
        }

        It 'calls OutTextBlankLine' {
            Mock -CommandName OutTextBlankLine -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { BlankLine } | OutTextSection;
            Assert-MockCalled -CommandName OutTextBlankLine -Exactly 1;
        }

        It 'warns on call OutTextTOC' {
            Mock -CommandName OutTextTOC -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { TOC 'TestTOC' } | OutTextSection -WarningAction SilentlyContinue;
            Assert-MockCalled OutTextTOC -Exactly 0;
        }

        It 'calls nested OutXmlSection' {
            ## Note this must be called last in the Describe script block as the OutXmlSection gets mocked!
            Mock -CommandName OutTextSection -Verifiable -MockWith { };
            Section -Name TestSection -ScriptBlock { Section -Name SubSection { } } | OutTextSection;
            Assert-MockCalled -CommandName OutTextSection -Exactly 1;
        }

    }

    Describe 'OutTextTable' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        Context 'As Table.' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120.' {
                $table = Table -Hashtable $services -Name 'Test Table' | OutTextTable;
                $table.Length | Should Be 212;
            }

            It 'Set width with of 35.' {
                $Options = New-PScriboTextOptions -TextWidth 35;
                $table = Table -Hashtable $services -Name 'Test Table' | OutTextTable;
                $table.Length | Should Be 335; ## Text tables are now set to wrap..
            }

        } #end context table

        Context 'As List.' {
        
            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120.' {
                $table = Table -Hashtable $services 'Test Table' -List | OutTextTable;
                $table.Length | Should Be 255;
            }

            It 'Default width of 25.' {
                $Options = New-PScriboTextOptions -TextWidth 25;
                $table = Table -Hashtable $services 'Test Table' -List | OutTextTable;
                $table.Length | Should Be 357;
            }

        } #end context table

    }

} #end inmodulescope

<#Code coverage report:
Covered 80.00% of 110 analyzed commands in 1 file.

Missed commands:

File                 Function         Line Command                                                                                                                                       
----                 --------         ---- -------                                                                                                                                       
OutText.Internal.ps1 OutTextTOC         49 $tocBuilder = New-Object -TypeName System.Text.StringBuilder                                                                                  
OutText.Internal.ps1 OutTextTOC         50 [ref] $null = $tocBuilder.AppendLine($TOC.Name)                                                                                               
OutText.Internal.ps1 OutTextTOC         51 [ref] $null = $tocBuilder.AppendLine(''.PadRight($Options.SeparatorWidth, $Options.SectionSeparator))                                         
OutText.Internal.ps1 OutTextTOC         52 $maxSectionNumberLength = ($Document.TOC.Number | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum).Length                     
OutText.Internal.ps1 OutTextTOC         52 $Document.TOC.Number                                                                                                                          
OutText.Internal.ps1 OutTextTOC         52 Measure-Object -Maximum                                                                                                                       
OutText.Internal.ps1 OutTextTOC         52 Select-Object -ExpandProperty Maximum                                                                                                         
OutText.Internal.ps1 OutTextTOC         53 $Document.TOC                                                                                                                                 
OutText.Internal.ps1 OutTextTOC         54 $sectionNumberPaddingLength = $maxSectionNumberLength - $tocEntry.Number.Length                                                               
OutText.Internal.ps1 OutTextTOC         55 $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ')                                                                                      
OutText.Internal.ps1 OutTextTOC         56 $sectionPadding = ''.PadRight($sectionNumberPaddingLength, ' ')                                                                               
OutText.Internal.ps1 OutTextTOC         57 [ref] $null = $tocBuilder.AppendFormat('{0}{1}  {2}{3}', $tocEntry.Number, $sectionPadding, $sectionNumberIndent, $tocEntry.Name).AppendLine()
OutText.Internal.ps1 OutTextTOC         59 return $tocBuilder.ToString()                                                                                                                 
OutText.Internal.ps1 OutTextSection     94 [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name                                                                           
OutText.Internal.ps1 OutTextSection    100 $sectionId = '{0}..' -f $s.Id.Substring(0,38)                                                                                                 
OutText.Internal.ps1 OutTextSection    104 [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s))                                                                            
OutText.Internal.ps1 OutTextSection    104 OutTextSection -Section $s                                                                                                                    
OutText.Internal.ps1 OutTextParagraph  129 $text = "$padding$($Paragraph.Text)"                                                                                                          
OutText.Internal.ps1 OutTextParagraph  129 $Paragraph.Text                                                                                                                               
OutText.Internal.ps1 OutTextLineBreak  145 $Options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1                                                                                       
OutText.Internal.ps1 OutTextTable      174 $Options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1                                                                                       
OutText.Internal.ps1 OutStringWrap     212 $textBuilder = $null                                        
#>
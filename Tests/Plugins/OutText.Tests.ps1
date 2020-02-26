$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    $isNix = $false
    if (($PSVersionTable['PSEdition'] -eq 'Core') -and (-not $IsWindows))
    {
        $isNix = $true
    }

    Describe 'OutText\OutText' {
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

            Document -Name 'TestDocument' -ScriptBlock { Get-Process | Select-Object -First 1 | Table 'TestTable' } | OutText -Path $path;

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

    Describe 'OutText.Internal\OutTextBlankLine' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        It 'Defaults to a single blank line.' {
            $expected = [System.Environment]::NewLine

            $l = BlankLine | OutTextBlankLine;

            $l | Should Be $expected
        }

        It 'Creates 3 blank lines.' {
            $expected = '{0}{0}{0}' -f [System.Environment]::NewLine

            $l = BlankLine -Count 3 | OutTextBlankLine;

            $l | Should Be $expected
        }

    }

    Describe 'OutText.Internal\OutTextLineBreak' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
        $Options = New-PScriboTextOption;

        It 'Defaults to 120 and includes new line.' {
            $expected = 122
            if ($isNix) { $expected -= 1 }

            $l = OutTextLineBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes new line.' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 40
            $expected = 42
            if ($isNix) { $expected -= 1 }

            $l = OutTextLineBreak;

            $l.Length | Should Be $expected;
        }

        It 'Wraps lines and includes new line' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 80;
            $expected = 84
            if ($isNix) { $expected -= 2 }

            $l = OutTextLineBreak

            $l.Length | Should Be $expected;
        }

    } #end describe OutTextLineBreak

    Describe 'OutText.Internal\OutTextPageBreak' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
        $Options = New-PScriboTextOption;

        It 'Defaults to 120 and includes new line .' {
            $expected = 124
            if ($isNix) { $expected -= 2 }

            $l = OutTextPageBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes new line.' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 40
            $expected = 44
            if ($isNix) { $expected -= 2 }

            $l = OutTextPageBreak

            $l.Length | Should Be $expected
        }

        It 'Wraps lines and includes new line.' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 80
            $expected = 86
            if ($isNix) { $expected -= 3 }

            $l = OutTextPageBreak

            $l.Length | Should Be $expected
        }

    } #end describe OutTextLineBreak

    Describe 'OutText.Internal\OutTextParagraph' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};
        $Options = New-PScriboTextOption;

        Context 'By pipeline.' {

            It 'Paragraph with new line.' {
                $testParagraph = 'Test paragraph.';
                $expected = 'Test paragraph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | OutTextParagraph;

                $p | Should BeExactly $expected;
            }

            It 'Paragraph with no new line.' {
                $testParagraph = 'Test paragraph.';

                $p = Paragraph $testParagraph -NoNewLine | OutTextParagraph;

                $p | Should BeExactly $testParagraph;
            }

            It 'Paragraph wraps at 10 characters with new line.' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Test paragraph.'
                $expected = 'Test parag{0}raph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | OutTextParagraph

                $p | Should BeExactly $expected
            }

             It 'Paragraph wraps at 10 characters with no new line.' {
                $testParagraph = 'Test paragraph.'
                $Options = New-PScriboTextOption -TextWidth 10
                $expected = 'Test parag{0}raph.' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph -NoNewLine | OutTextParagraph;

                $p | Should BeExactly $expected;
            }

        } #end context by pipeline

        Context 'By named -Paragraph parameter.' {

            It 'By named -Paragraph parameter with new line.' {
                $testParagraph = 'Test paragraph.';
                $expected = '{0}{1}' -f $testParagraph, [System.Environment]::NewLine

                $p = OutTextParagraph -Paragraph (Paragraph $testParagraph)

                $p | Should BeExactly $expected
            }

            It 'By named -Paragraph parameter with no new line.' {
                $testParagraph = 'Test paragraph.';

                $p = OutTextParagraph -Paragraph (Paragraph $testParagraph -NoNewLine);

                $p | Should BeExactly $testParagraph;
            }
        } #end context -paragraph

    } #end describe outtextparagraph

    Describe 'OutText.Internal\OutTextSection' {
        $Document = Document -Name 'TestDocument' -ScriptBlock { };
        $pscriboDocument = $Document;

        It 'outputs indented section (#73)' {
            $result = Section -Name TestSection -ScriptBlock { } -Tabs 2 | OutTextSection;

            $result -match '^\r?\n        ' | Should Be $true;
        }

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

            Section -Name TestSection -ScriptBlock { Get-Process | Select-Object -First 3 | Table TestTable } | OutTextSection;

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

    Describe 'OutText.Internal\OutTextTable' {
        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {};

        Context 'As Table.' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120.' {
                $expected = 208
                if ($isNix) { $expected -= 6 }

                $table = Table -Hashtable $services -Name 'Test Table' | OutTextTable;

                $table.Length | Should Be $expected;  # Trailing spaces are removed (#67)
            }

            It 'Set width with of 35.' {
                $Options = New-PScriboTextOption -TextWidth 35;
                $expected = 313
                if ($isNix) { $expected -= 10 }

                $table = Table -Hashtable $services -Name 'Test Table' | OutTextTable;

                $table.Length | Should Be $expected; ## Text tables are now set to wrap.. Trailing spaces are removed (#67)
            }

        } #end context table

        Context 'As List.' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120.' {
                $expected = 255
                if ($isNix) { $expected -= 12 }

                $table = Table -Hashtable $services 'Test Table' -List | OutTextTable;

                $table.Length | Should Be $expected;
            }

            It 'Default width of 25.' {
                $Options = New-PScriboTextOption -TextWidth 25;
                $expected = 354
                if ($isNix) { $expected -= 18 }

                $table = Table -Hashtable $services 'Test Table' -List | OutTextTable;

                $table.Length | Should Be $expected; # Trailing spaces are removed (#67)
            }

        } #end context table

    }

    Describe 'OutText.Internal\OutTextTOC' {

        It 'outputs TOC name and section break' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1';
            $heading2 = 'Heading 2';
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{1}-+' -f $tocName, [System.Environment]::NewLine;

            $result = OutTextTOC -TOC $Document.Sections[0];

            $result | Should MatchExactly $expected;
        }

        It 'adds section numbers when "EnableSectionNumbering" is enabled (#20)' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1';
            $heading2 = 'Heading 2';
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                DocumentOption -EnableSectionNumbering
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{3}-+{3}1\s+{1}{3}1.1\s+{2}{3}$' -f $tocName, $heading1, $heading2, [System.Environment]::NewLine

            $options = Merge-PScriboPluginOption -DocumentOptions $Document.Options -PluginOptions (New-PScriboTextOption)
            $result = OutTextTOC -TOC $Document.Sections[0] -Verbose;

            $result | Should Match $expected;
        }

        It 'does not add section numbers when "EnableSectionNumbering" is disabled (#20)' {

            $tocName = 'Table of contents'
            $heading1 = 'Heading 1';
            $heading2 = 'Heading 2';
            $Document  = Document -Name 'TestDocument' -ScriptBlock {
                TOC -Name $tocName
                Section $heading1 -Style Heading1 {
                    Section $heading2 -Style Heading2 { }
                }
            }
            $expected = '^{0}{3}-+{3}{1}{3} {2}{3}$' -f $tocName, $heading1, $heading2, [System.Environment]::NewLine

            $result = OutTextTOC -TOC $Document.Sections[0] -Verbose;

            $result | Should Match $expected;
        }
    }

    Describe 'OutText.Internal\OutTextImage' {

        It 'adds alttext to image' {
            $testImage = [PSCustomObject] @{
                Text     = 'Dummy Image'
            }
            $expected = '\[Image Text="{0}"\]' -f $testImage.Text;

            $result = OutTextImage -Image $testImage

            $result -match $expected | Should Be $true
        }

    }

} #end inmodulescope

<#
Code coverage report:
Covered 75.00% of 164 analyzed commands in 1 file.

Missed commands:

File                 Function         Line Command
----                 --------         ---- -------
OutText.Internal.ps1 OutTextTOC         74 $options = Get-Variable -Name Options -ValueOnly
OutText.Internal.ps1 OutTextTOC         75 if (-not ($options.ContainsKey('SeparatorWidth'))) {...
OutText.Internal.ps1 OutTextTOC         75 $options.ContainsKey('SeparatorWidth')
OutText.Internal.ps1 OutTextTOC         76 $options['SeparatorWidth'] = 120
OutText.Internal.ps1 OutTextTOC         78 if (-not ($options.ContainsKey('LineBreakSeparator'))) {...
OutText.Internal.ps1 OutTextTOC         78 $options.ContainsKey('LineBreakSeparator')
OutText.Internal.ps1 OutTextTOC         79 $options['LineBreakSeparator'] = '_'
OutText.Internal.ps1 OutTextTOC         81 if (-not ($options.ContainsKey('TextWidth'))) {...
OutText.Internal.ps1 OutTextTOC         81 $options.ContainsKey('TextWidth')
OutText.Internal.ps1 OutTextTOC         82 $options['TextWidth'] = 120
OutText.Internal.ps1 OutTextTOC         84 if(-not ($Options.ContainsKey('SectionSeparator'))) {...
OutText.Internal.ps1 OutTextTOC         84 $Options.ContainsKey('SectionSeparator')
OutText.Internal.ps1 OutTextTOC         85 $options['SectionSeparator'] = "-"
OutText.Internal.ps1 OutTextSection    147 $options = Get-Variable -Name Options -ValueOnly
OutText.Internal.ps1 OutTextSection    148 if (-not ($options.ContainsKey('SeparatorWidth'))) {...
OutText.Internal.ps1 OutTextSection    148 $options.ContainsKey('SeparatorWidth')
OutText.Internal.ps1 OutTextSection    149 $options['SeparatorWidth'] = 120
OutText.Internal.ps1 OutTextSection    151 if (-not ($options.ContainsKey('LineBreakSeparator'))) {...
OutText.Internal.ps1 OutTextSection    151 $options.ContainsKey('LineBreakSeparator')
OutText.Internal.ps1 OutTextSection    152 $options['LineBreakSeparator'] = '_'
OutText.Internal.ps1 OutTextSection    154 if (-not ($options.ContainsKey('TextWidth'))) {...
OutText.Internal.ps1 OutTextSection    154 $options.ContainsKey('TextWidth')
OutText.Internal.ps1 OutTextSection    155 $options['TextWidth'] = 120
OutText.Internal.ps1 OutTextSection    157 if(-not ($Options.ContainsKey('SectionSeparator'))) {...
OutText.Internal.ps1 OutTextSection    157 $Options.ContainsKey('SectionSeparator')
OutText.Internal.ps1 OutTextSection    158 $options['SectionSeparator'] = "-"
OutText.Internal.ps1 OutTextSection    167 [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name
OutText.Internal.ps1 OutTextSection    173 $sectionId = '{0}..' -f $s.Id.Substring(0,38)
OutText.Internal.ps1 OutTextSection    176 $currentIndentationLevel = $s.Level +1
OutText.Internal.ps1 OutTextSection    179 [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s))
OutText.Internal.ps1 OutTextSection    179 OutTextSection -Section $s
OutText.Internal.ps1 OutTextParagraph  211 $options['TextWidth'] = 120
OutText.Internal.ps1 OutTextParagraph  221 $text = "$padding$($Paragraph.Text)"
OutText.Internal.ps1 OutTextParagraph  221 $Paragraph.Text
OutText.Internal.ps1 OutTextLineBreak  246 $options['SeparatorWidth'] = 120
OutText.Internal.ps1 OutTextLineBreak  249 $options['LineBreakSeparator'] = '_'
OutText.Internal.ps1 OutTextLineBreak  252 $options['TextWidth'] = 120
OutText.Internal.ps1 OutTextLineBreak  259 $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1
OutText.Internal.ps1 OutTextTable      300 $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1
OutText.Internal.ps1 OutStringWrap     334 $Width = 4096
OutText.Internal.ps1 OutStringWrap     348 $textBuilder = $null
#>

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Section\Section' {
        $pscriboDocument = Document 'ScaffoldDocument' {};

        Context 'By named parameters' {

            It 'calls a nested element.' {
                # function PageBreak { }
                Mock PageBreak -Verifiable { }
                Section -Name 'Test Section' -ScriptBlock { PageBreak };
                Assert-VerifiableMock;
            }

            It 'returns a PSCustomObject object.' {
                $section = Section -Name 'Test Section' -ScriptBlock { };
                $section.GetType().Name | Should Be 'PSCustomObject';
                $section.Type | Should Be 'PScribo.Section';
            }

            It 'creates an array list of sections.' {
                $section = Section -Name 'Test Section' -ScriptBlock { };
                $section.Sections.GetType() | Should BeExactly 'System.Collections.ArrayList';
            }

            It 'creates an empty array list of sections.' {
                $section = Section -Name 'Test Section' -ScriptBlock { };
                $section.Sections.Count | Should Be 0;
            }

            It 'creates a concatenated section Id.' {
                $section = Section -Name 'Test Section' -ScriptBlock { };
                $section.Id | Should BeExactly 'TESTSECTION';
            }

            It 'creates a section by named -Name and -ScriptBlock parameters.' {
                $section = Section -Name 'Test Section' -ScriptBlock { };
                $section.Name | Should BeExactly 'Test Section';
            }

            It 'creates a section by named -Name, -Style and -ScriptBlock parameters.' {
                $section = Section -Name 'Test Section' -Style 'Test Style' -ScriptBlock { };
                $section.Style | Should BeExactly 'Test Style';
            }

            It 'creates a section by named -ExcludeFromTOC and -ScriptBlock parameters.' {
                $section = Section -Name 'Test Section' -ExcludeFromTOC -ScriptBlock { };
                 $section.IsExcluded | Should Be $true;
            }

            It 'creates a section by named -Name, -ExcludeFromTOC and -ScriptBlock parameters.' {
                $section = Section -Name 'Test Section' -ExcludeFromTOC -ScriptBlock { };
                $section.IsExcluded | Should Be $true;
            }

        } #end context By named parameters

        Context 'By positional parameters' {
            It 'creates a section by positional -Name and -ScriptBlock parameters.' {
                $section = Section 'Test Section' { };
                $section.Name | Should BeExactly 'Test Section';
            }

            It 'creates a section by positional -Name and -ScriptBlock parameters and named -Style parameter.' {
                $section = Section 'Test Section' -Style 'Test Style' { };
                $section.Style | Should BeExactly 'Test Style';
            }

            It 'creates a section by positional -Name and -ScriptBlock parameters and named -ExcludeFromTOC parameter.' {
                $section = Section 'Test Section' -ExcludeFromTOC { };
                $section.IsExcluded | Should Be $true;
            }

            It 'creates a section by positional -Name and -ScriptBlock parameters and named -Style and -ExcludeFromTOC parameter.' {
                $section = Section 'Test Section' -Style 'Test Style' -ExcludeFromTOC { };
                $section.IsExcluded | Should Be $true;
                $section.Style | Should BeExactly 'Test Style';
            }

        } #end context By positional parameters

    }

} #end inmodulescope

<#
Missed commands:

File        Function Line Command
----        -------- ---- -------
Section.ps1 Section    28 [ref] $null = $pscriboSection.Sections.Add($result)
#>

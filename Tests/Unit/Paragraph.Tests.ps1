$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Paragraph\Paragraph' {

        Context 'By Named Parameter' {
            $pscriboDocument = Document 'ScaffoldDocument' {};

            It 'returns a PSCustomObject object.' {
                $p = Paragraph -Name Test;
                $p.GetType().Name | Should Be 'PSCustomObject';
            }

            It 'creates a PScribo.Paragraph type.' {
                $p = Paragraph -Name Test;
                $p.Type | Should Be 'PScribo.Paragraph';
            }

            It 'creates paragraph by named -Name parameter.' {
                $text = 'Simple paragraph.';
                $p = Paragraph -Name $text;
                $p.Id | Should Be $text;
            }

            It 'creates paragraph by named -Name and -Text parameters.' {
                $text = 'Simple paragraph.';
                $p = Paragraph -Name Test -Text $text;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
            }

            It 'creates paragraph by named -Name, -Text and -Style parameters.' {
                $text = 'Simple paragraph.';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Style $style;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
                $p.Style | Should Be $style;
            }

            It 'creates paragraph by named -Name, -Text and -Value parameters.' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
                $p.Value | Should Be $value;
            }

            It 'creates paragraph by named -Name, -Text, -Value and -Style parameters.' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
                $p.Value | Should Be $value;
                $p.Style | Should Be $style;
            }

            It 'creates a paragraph with custom Bold formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Bold;
                $p.Bold | Should Be $true;
            }

            It 'creates a paragraph with custom Italic formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Italic;
                $p.Italic | Should Be $true;
            }

            It 'creates a paragraph with custom Underline formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Underline;
                $p.Underline | Should Be $true;
            }

            It 'creates a paragraph with custom Size formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Size 14;
                $p.Size | Should Be 14;
            }

            It 'creates a paragraph with custom Color formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Color ff0;
                $p.Color | Should Be 'ff0';
            }

            It 'creates a paragraph with custom Font formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Font 'Courier New';
                $p.Font | Should Be 'Courier New';
            }

            It 'creates a paragraph with custom Font[] formatting' {
                $text = 'Simple paragraph.';
                $value = 'Test';
                $style = 'Test';
                $p = Paragraph -Name Test -Text $text -Value $value -Style $style -Font 'Courier New','Consolas';
                $p.Font -contains 'Courier New' | Should Be $true;
                $p.Font -contains 'Consolas' | Should Be $true;
            }

        } #end context By Named Parameters

        Context 'By Positional Parameters' {
            $pscriboDocument = Document 'ScaffoldDocument' {};

            It 'creates paragraph by positional -Name parameter.' {
                $text = 'Simple paragraph.';
                $p = Paragraph $text;
                $p.Id | Should Be $text;
            }

            It 'creates paragraph by positional -Name and -Text parameters.' {
                $text = 'Simple paragraph.';
                $p = Paragraph Test $text;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
            }

            It 'creates paragraph by positional -Name, -Text and -Value parameters.' {
                $text = 'Simple paragraph.';
                $value = 'TestValue';
                $p = Paragraph Test $text $value;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
                $p.Value | Should Be $value;
            }

            It 'creates paragraph by positional -Name, -Text and -Value and named -Style parameters.' {
                $text = 'Simple paragraph.';
                $value = 'TestValue';
                $style = 'Test';
                $p = Paragraph Test $text $value -Style $style;
                $p.Id | Should Be 'Test';
                $p.Text | Should Be $text;
                $p.Value | Should Be $value;
                $p.Style | Should Be $style;
            }
        } #end context By Positional Parameters

    } #end describe Paragraph

} #end inmodulescope
<#
Missed commands:

File          Function  Line Command
----          --------  ---- -------
Paragraph.ps1 Paragraph   27 $paragraphDisplayName = '{0}..' -f $Name.Substring(18)
#>

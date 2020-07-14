$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Paragraph' {

        Context 'By Named Parameter' {

            $pscriboDocument = Document 'ScaffoldDocument' {}

            It 'returns a PSCustomObject object' {
                $p = Paragraph -Name Test

                $p.GetType().Name | Should Be 'PSCustomObject'
            }

            It 'creates a PScribo.Paragraph type' {
                $p = Paragraph -Name Test

                $p.Type | Should Be 'PScribo.Paragraph'
            }

            It 'creates paragraph by named -Name parameter' {
                $text = 'Simple paragraph.'

                $p = Paragraph -Name $text

                $p.Id | Should Be $text
            }

            It 'creates paragraph by named -Name and -Text parameters' {
                $text = 'Simple paragraph.'

                $p = Paragraph -Name Test -Text $text

                $p.Id | Should Be 'Test'
                $p.Sections[0].Text | Should Be $text
            }

            It 'creates paragraph by named -Name, -Text and -Style parameters' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style

                $p.Id | Should Be 'Test'
                $p.Sections[0].Text | Should Be $text
                $p.Style | Should Be $style
            }

            It 'creates a paragraph with custom Bold formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Bold

                $p.Sections[0].Bold | Should Be $true
            }

            It 'creates a paragraph with custom Italic formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Italic

                $p.Sections[0].Italic | Should Be $true
            }

            It 'creates a paragraph with custom Underline formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Underline

                $p.Sections[0].Underline | Should Be $true
            }

            It 'creates a paragraph with custom Size formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Size 14

                $p.Sections[0].Size | Should Be 14
            }

            It 'creates a paragraph with custom Color formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Color ff0

                $p.Sections[0].Color | Should Be 'ff0'
            }

            It 'creates a paragraph with custom Font formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Font 'Courier New'

                $p.Sections[0].Font | Should Be 'Courier New'
            }

            It 'creates a paragraph with custom Font[] formatting' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph -Name Test -Text $text -Style $style -Font 'Courier New','Consolas'

                $p.Sections[0].Font -contains 'Courier New' | Should Be $true
                $p.Sections[0].Font -contains 'Consolas' | Should Be $true
            }
        }

        Context 'By Positional Parameters' {

            $pscriboDocument = Document 'ScaffoldDocument' {}

            It 'creates paragraph by positional -Name parameter' {
                $text = 'Simple paragraph.'

                $p = Paragraph $text

                $p.Id | Should Be $text
            }

            It 'creates paragraph by positional -Name and -Text parameters' {
                $text = 'Simple paragraph.'

                $p = Paragraph Test $text

                $p.Id | Should Be 'Test'
                $p.Sections[0].Text | Should Be $text
            }

            It 'creates paragraph by positional -Name, -Text and named -Style parameters' {
                $text = 'Simple paragraph.'
                $style = 'Test'

                $p = Paragraph Test $text -Style $style

                $p.Id | Should Be 'Test'
                $p.Sections[0].Text | Should Be $text
                $p.Style | Should Be $style
            }
        }
    }
}

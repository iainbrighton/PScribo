$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextParagraph' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {}
        $Options = New-PScriboTextOption

        Context 'By pipeline' {

            It 'Paragraph with new line' {
                $testParagraph = 'Test paragraph.'
                $expected = 'Test paragraph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'Paragraph with no new line' {
                $testParagraph = 'Test paragraph.'

                $p = Paragraph $testParagraph -NoNewLine | Out-TextParagraph

                $p | Should BeExactly $testParagraph
            }

            It 'Paragraph wraps at 10 characters with new line' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Test paragraph.'
                $expected = 'Test parag{0}raph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'Paragraph wraps at 10 characters with no new line' {
                $testParagraph = 'Test paragraph.'
                $Options = New-PScriboTextOption -TextWidth 10
                $expected = 'Test parag{0}raph.' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph -NoNewLine | Out-TextParagraph

                $p | Should BeExactly $expected
            }
        }

        Context 'By named -Paragraph parameter' {

            It 'By named -Paragraph parameter with new line' {
                $testParagraph = 'Test paragraph.'
                $expected = '{0}{1}' -f $testParagraph, [System.Environment]::NewLine

                $p = Out-TextParagraph -Paragraph (Paragraph $testParagraph)

                $p | Should BeExactly $expected
            }

            It 'By named -Paragraph parameter with no new line' {
                $testParagraph = 'Test paragraph.'

                $p = Out-TextParagraph -Paragraph (Paragraph $testParagraph -NoNewLine)

                $p | Should BeExactly $testParagraph
            }
        }
    }
}

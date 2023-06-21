$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Json\Out-JsonParagraph' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $pscriboDocument = $Document
        $Options = New-PScriboJsonOption

        Context 'By pipeline' {

            It 'Paragraph outputs text' {
                $testParagraph = 'Test paragraph.'
                $expected = 'Test paragraph.'

                $p = Paragraph $testParagraph | Out-JsonParagraph

                $p | Should BeExactly $expected
            }

            It 'adds spaces between text runs (by default)' {
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = 'Test paragraph'

                $p = Paragraph $testParagraph | Out-JsonParagraph

                $p | Should BeExactly $expected
            }

            It 'does not add spaces between text runs (when specified)' {
                $testParagraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }
                $expected = 'Testparagraph'

                $p = Paragraph $testParagraph | Out-JsonParagraph

                $p | Should BeExactly $expected
            }
        }

        Context 'By named -Paragraph parameter' {

            It 'By named -Paragraph parameter' {
                $testParagraph = 'Test paragraph.'
                $expected = 'Test paragraph.'

                $p = Out-JsonParagraph -Paragraph (Paragraph $testParagraph)

                $p | Should BeExactly $expected
            }

        }
    }
}

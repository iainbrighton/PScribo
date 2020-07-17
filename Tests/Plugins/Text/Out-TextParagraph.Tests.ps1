$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextParagraph' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $pscriboDocument = $Document
        $Options = New-PScriboTextOption

        Context 'By pipeline' {

            It 'Paragraph with new line' {
                $testParagraph = 'Test paragraph.'
                $expected = 'Test paragraph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'Paragraph wraps at 10 characters with new line' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Testparagraph.'
                $expected = 'Testparagr{0}aph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'Paragraph breaks on space with new line' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Test paragraph.'
                $expected = 'Test{0}paragraph.{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'adds spaces between text runs (by default)' {
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = 'Test paragraph{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

                $p | Should BeExactly $expected
            }

            It 'does not add spaces between text runs (when specified)' {
                $testParagraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }
                $expected = 'Testparagraph{0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-TextParagraph

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

        }
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Markdown\Out-MarkdownParagraph' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $pscriboDocument = $Document
        $Options = New-PScriboMarkdownOption

        Context 'By pipeline' {

            It 'ends with "  " and a new line' {
                $testParagraph = 'Test paragraph.'
                $expected = 'Test paragraph.  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'wraps at 10 characters and ends with "  " and a new line' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Testparagraph.'
                $expected = 'Testparagr  {0}aph.  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'wraps on space and ends with "  " and a new line' {
                $Options = New-PScriboTextOption -TextWidth 10
                $testParagraph = 'Test paragraph.'
                $expected = 'Test{0}paragraph.  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'preserves line breaks with "  " and a new line' {
                $Options = New-PScriboTextOption
                $testParagraph = "Test`n`nparagraph."
                $expected = 'Test  {0}paragraph.  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'adds spaces between text runs (by default)' {
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = 'Test paragraph  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'does not add spaces between text runs (when specified)' {
                $testParagraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }
                $expected = 'Testparagraph  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph with paragraph bold style' {
                $customStyle = Style -Name 'Custom' -Bold
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = '**Test paragraph**  {0}' -f [System.Environment]::NewLine

                $p = Paragraph -Style 'Custom' -ScriptBlock $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph with paragraph italic style' {
                $customStyle = Style -Name 'Custom' -Italic
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = '_Test paragraph_  {0}' -f [System.Environment]::NewLine

                $p = Paragraph -Style 'Custom' -ScriptBlock $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph with paragraph bold and italic style' {
                $customStyle = Style -Name 'Custom' -Bold -Italic
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = '***Test paragraph***  {0}' -f [System.Environment]::NewLine

                $p = Paragraph -Style 'Custom' -ScriptBlock $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with run bold style' {
                $customStyle = Style -Name 'Custom' -Bold
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph' -Style 'Custom'
                }
                $expected = 'Test **paragraph**  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with run italic style' {
                $customStyle = Style -Name 'Custom' -Italic
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph' -Style 'Custom'
                }
                $expected = 'Test _paragraph_  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with run bold and italic style' {
                $customStyle = Style -Name 'Custom' -Bold -Italic
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph' -Style 'Custom'
                }
                $expected = 'Test ***paragraph***  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with inline bold style' {
                $testParagraph = {
                    Text 'Test' -Bold
                    Text 'paragraph'
                }
                $expected = '**Test **paragraph  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with inline italic style' {
                $testParagraph = {
                    Text 'Test' -Italic
                    Text 'paragraph'
                }
                $expected = '_Test _paragraph  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }

            It 'formats paragraph run with inline bold and italic style' {
                $testParagraph = {
                    Text 'Test'
                    Text 'paragraph' -Bold -Italic
                }
                $expected = 'Test ***paragraph***  {0}' -f [System.Environment]::NewLine

                $p = Paragraph $testParagraph | Out-MarkdownParagraph

                $p | Should BeExactly $expected
            }
        }

        Context 'By named -Paragraph parameter' {

            It 'By named -Paragraph parameter ends with "  " and a new line' {
                $testParagraph = 'Test paragraph.'
                $expected = '{0}  {1}' -f $testParagraph, [System.Environment]::NewLine

                $p = Out-MarkdownParagraph -Paragraph (Paragraph $testParagraph)

                $p | Should BeExactly $expected
            }

        }
    }
}

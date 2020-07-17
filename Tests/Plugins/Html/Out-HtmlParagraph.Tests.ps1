$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlParagraph' {

        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { }
        $Document = $pscriboDocument
        $script:currentPageNumber = 1

        Context 'By Named Parameter' {

            It 'creates paragraph with no style and new line' {
                Paragraph 'Test paragraph.' | Out-HtmlParagraph | Should BeExactly "<div>Test paragraph.</div>"
            }

            It 'creates paragraph with custom name/id' {
                Paragraph -Name 'Test' -Text 'Test paragraph.' | Out-HtmlParagraph | Should BeExactly "<div>Test paragraph.</div>"
            }

            It 'creates paragraph with named -Style parameter' {
                Paragraph 'Test paragraph.' -Style Named | Out-HtmlParagraph | Should BeExactly "<div class=`"Named`">Test paragraph.</div>"
            }

            It 'encodes HTML paragraph content' {
                $expected = '<div>Embedded &lt;br /&gt;</div>'

                $result = Paragraph 'Embedded <br />' | Out-HtmlParagraph

                $result | Should BeExactly $expected
            }

            It 'creates paragraph with embedded new line' {
                $paragraphText = 'Embedded{0}New Line' -f [System.Environment]::NewLine
                $expected = '<div>Embedded<br />New Line</div>'

                $result = Paragraph $paragraphText | Out-HtmlParagraph

                $result | Should BeExactly $expected
            }

            It 'adds spaces between text runs (by default)' {
                $paragraph = {
                    Text 'Test'
                    Text 'paragraph'
                }
                $expected = '<div>Test paragraph</div>'

                $result = Paragraph $paragraph | Out-HtmlParagraph

                $result | Should BeExactly $expected
            }

            It 'does not add spaces between text runs (when specified)' {
                $paragraph = {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }
                $expected = '<div>Testparagraph</div>'

                $result = Paragraph $paragraph | Out-HtmlParagraph

                $result | Should BeExactly $expected
            }

        } #end context By Named Parameter

        It 'uses invariant culture paragraph size (#6)' {
            $currentCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture.Name
            [System.Threading.Thread]::CurrentThread.CurrentCulture = 'da-DK'

            $null = (Paragraph 'Test paragraph.' -Size 11 | Out-HtmlParagraph) -match '(?<=style=").+(?=">)'
            $fontSize = ($Matches[0]).Trim(';')

            ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92rem'
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture
        }

    }
}

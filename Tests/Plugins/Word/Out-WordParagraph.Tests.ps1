$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginsRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginsRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    function GetMatch
    {
        [CmdletBinding()]
        param
        (
            [System.String] $String,
            [System.Management.Automation.SwitchParameter] $Complete
        )
        Write-Verbose "Pre Match : '$String'"
        $matchString = $String.Replace('/','\/')
        if (-not $String.StartsWith('^'))
        {
            $matchString = $matchString.Replace('[..]','[\s\S]+')
            $matchString = $matchString.Replace('[??]','([\s\S]+)?')
            if ($Complete)
            {
                $matchString = '^<w:test xmlns:w="http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main">{0}<\/w:test>$' -f $matchString
            }
        }
        Write-Verbose "Post Match: '$matchString'"
        return $matchString
    }

    Describe 'Plugins\Word\Out-WordParagraph' {

        It 'outputs paragraph "<w:p>[..]></w:p>"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph'
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p>[..]></w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs paragraph properties "<w:p><w:pPr>[..]></w:pPr[..]></w:p>"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph'
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p><w:pPr>[..]></w:pPr[..]></w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs indented paragraph "<w:p><w:pPr><w:ind w:left="1440" />[..]</w:p>"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Tabs 2
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p><w:pPr><w:ind w:left="1440" />[..]</w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs paragraph style "<w:p><w:pPr><w:pStyle w:val="[..]" />[..]</w:p>"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Style Heading3
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p><w:pPr><w:pStyle w:val="Heading3" />[..]</w:p>'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run "[..]<w:r>[..]></w:r>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Style Heading3
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:r>[..]></w:r>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs empty run properties "[..]<w:r><w:rPr />[..]></w:r>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Style Heading3
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:r><w:rPr />[..]></w:r>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property font "[..]<w:rPr><w:rFonts w:ascii="[..]" w:hAnsi="[..]" /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Font Ariel
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:rFonts w:ascii="Ariel" w:hAnsi="Ariel" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property font size "[..]<w:rPr><w:sz w:val="[..]" /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Size 10
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:sz w:val="20" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property bold "[..]<w:rPr><w:b /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Bold
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:b /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property italic "[..]<w:rPr><w:i /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Italic
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:i /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property underline "[..]<w:rPr><w:u w:val="single" /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Underline
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:u w:val="single" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run property colour "[..]<w:rPr><w:color w:val="112233" /></w:rPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Paragraph 'Test paragraph' -Color 123
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:rPr><w:color w:val="112233" /></w:rPr>[..]'
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run text "[..]<w:r>[..]<w:t>{0}</w:t>[..]" using "Name" property' {
            $testParagraphText = 'Test paragraph'
            $document = Document -Name 'TestDocument' {
                Paragraph $testParagraphText -Font Ariel
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch ('[..]<w:r>[..]<w:t>{0}</w:t>[..]' -f $testParagraphText);
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run text "[..]<w:r>[..]<w:t>{0}</w:t>[..]" using "Text" property' {
            ## Ignore the space preservation namespace
            $testParagraphText = 'Test paragraph'
            $document = Document -Name 'TestDocument' {
                Paragraph -Name 'Test' -Text $testParagraphText
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch ('[..]<w:r>[..]<w:t>{0}</w:t>[..]' -f $testParagraphText);
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs run text "[..]<w:r>[..]<w:t></w:t><w:br /></w:t>[..]</w:r>[..]" with embedded new line' {
            ## Ignore the space preservation namespace
            $document = Document -Name 'TestDocument' {
                Paragraph -Name 'Test' -Text ('Test{0}Paragraph' -f [System.Environment]::NewLine)
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch ('<w:t>Test</w:t><w:br /><w:t>Paragraph</w:t></w:r>');
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'outputs space between runs "<w:t [..]>abc </w:t>[..]<w:t>def</w:t>" (by default)' {
            ## Ignore the space preservation namespace
            $document = Document -Name 'TestDocument' {
                Paragraph {
                    Text 'Test'
                    Text 'paragraph'
                }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch ('<w:t [..]>Test </w:t>[..]<w:t>paragraph</w:t>');
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

        It 'does not output space between runs "<w:t>abc</w:t>[..]<w:t>def</w:t>" (when specified)' {
            ## Ignore the space preservation namespace
            $document = Document -Name 'TestDocument' {
                Paragraph {
                    Text 'Test' -NoSpace
                    Text 'paragraph'
                }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch ('<w:t>Test</w:t>[..]<w:t>paragraph</w:t>');
            $testDocument.DocumentElement.OuterXml  | Should Match $expected;
        }

    }
}

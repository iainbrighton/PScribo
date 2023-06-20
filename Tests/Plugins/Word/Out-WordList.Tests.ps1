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

    Describe 'Plugins\Word\Out-WordList' {

        Context 'single-level list' {

            It 'outputs number properties "<w:p><w:pPr><w:numPr>[..]></w:numPr></w:pPr>[..]</w:p>" per item' {
                $document = Document -Name 'TestDocument' {
                                List -Item 'Apples', 'Oranges', 'Pears'
                            }
                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:p><w:pPr><w:numPr>[..]></w:numPr></w:pPr>[..]</w:p>){3}'

                $testDocument.DocumentElement.OuterXml  | Should Match $expected
            }

            It 'outputs number property level "<w:ilvl w:val="0" />" per item' {
                $document = Document -Name 'TestDocument' {
                                List -Item 'Apples', 'Oranges', 'Pears'
                            }
                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:ilvl w:val="0" />.*){3}'

                $testDocument.DocumentElement.OuterXml  | Should Match $expected
            }

            It 'outputs run "<w:r><w:rPr /><w:t>[..]</w:t></w:r>" per item' {
                $document = Document -Name 'TestDocument' {
                                List -Item 'Apples', 'Oranges', 'Pears'
                            }
                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:r><w:rPr /><w:t>[..]</w:t></w:r>.*){3}'

                $testDocument.DocumentElement.OuterXml  | Should Match $expected
            }

        }

        Context 'multi-level list' {

            It 'outputs run "<w:r><w:rPr /><w:t>[..]</w:t></w:r>" per item' {
                $document = Document -Name 'TestDocument' {
                    List -Numbered {
                        Item 'Apples'
                        List -Numbered {
                            Item 'Braeburn'
                            Item 'Granny Smith'
                        }
                        Item 'Oranges'
                        Item 'Pears'
                    }
                }
                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:r><w:rPr /><w:t>[..]</w:t></w:r>.*){5}'

                $testDocument.DocumentElement.OuterXml  | Should Match $expected
            }

            It 'outputs number property level "<w:ilvl w:val="1" />" per nested item' {

                $document = Document -Name 'TestDocument' {
                    List -Numbered {
                        Item 'Apples'
                        List -Numbered {
                            Item 'Braeburn'
                            Item 'Granny Smith'
                        }
                        Item 'Oranges'
                        Item 'Pears'
                    }
                }
                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:ilvl w:val="1" />.*){2}'

                $testDocument.DocumentElement.OuterXml | Should Match $expected
            }

            It 'outputs custom numbered list format' {

                $customNumberFormat = 'xYz-%%%.'
                $indent = 1500
                $hanging = 200

                $document = Document -Name 'TestDocument' {
                    NumberStyle -Id 'CustomNumberStyle' -Custom $customNumberFormat -Indent $indent -Hanging $hanging -Align Left
                    List -Numbered -NumberStyle CustomNumberStyle -Item 'Apples','Bananas','Oranges'
                }
                $testNumberingDocument = Get-WordNumberingDocument -Lists $document.Lists

                $testNumberingDocument.DocumentElement.OuterXml | Should Match '<mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="001, 002, 003, ..." /></mc:Choice>'
                $testNumberingDocument.DocumentElement.OuterXml | Should Match '<w:lvl w:ilvl="0">'
                $testNumberingDocument.DocumentElement.OuterXml | Should Match '<w:start w:val="1" />'
                $testNumberingDocument.DocumentElement.OuterXml | Should Match ('<w:lvlText w:val="{0}" />' -f $customNumberFormat.Replace('%%%','%1'))
                $testNumberingDocument.DocumentElement.OuterXml | Should Match ('<w:pPr><w:ind w:left="{0}" w:hanging="{1}" /></w:pPr>' -f $indent, $hanging)
            }

        }

    }
}

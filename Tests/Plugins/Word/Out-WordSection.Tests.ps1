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

    Describe 'Plugins\Word\Out-WordSection' {

        It 'appends section "<w:p>[..]</w:p>"' {
            $document = Document -Name 'TestDocument' {
                Section -Name TestSection { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p>[..]</w:p>';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'appends section spacing "[..]<w:pPr><w:spacing w:before="160" w:after="160" /></w:pPr>[..]"' {
            $document = Document -Name 'TestDocument' {
                Section -Name TestSection { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:pPr><w:spacing w:before="160" w:after="160" /></w:pPr>[..]';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'appends section style "[..]<w:pStyle w:val="CustomStyle" />[..]"' {
            $document = Document -Name 'TestDocument' {
                Style -Name 'CustomStyle' -Color AAA
                Section -Name TestSection -Style 'CustomStyle' { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:pStyle w:val="CustomStyle" />[..]';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'outputs indented section "<w:p><w:pPr><w:ind w:left="1440" />[..]</w:p>" (#73)' {
            $document = Document -Name 'TestDocument' {
                Section -Name TestSection -Tabs 2 { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:pPr>[??]<w:ind w:left="1440" />[..]</w:pPr>[..]';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'increases section spacing between section levels' {
            $document = Document -Name 'TestDocument' {
                Section -Name SectionLevel0 {
                    Section -Name SectionLevel1 {
                        Section -Name SectionLevel2 { }
                    }
                }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:pPr><w:spacing w:before="240" w:after="240" /></w:pPr>[..]';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'appends "[..]<w:r><w:t>Section Run</w:t></w:r></w:p>" run' {
            $document = Document -Name 'TestDocument' {
                Section -Name 'Section Run' { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:r><w:t>Section Run</w:t></w:r></w:p>';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }

        It 'adds section numbering when enabled' {
            $document = Document -Name 'TestDocument' {
                DocumentOption -EnableSectionNumbering
                Section -Name 'SectionLevel1' { }
                Section -Name 'Numbered Section' { }
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '[..]<w:r><w:t>2 Numbered Section</w:t></w:r></w:p>';
            $testDocument.DocumentElement.OuterXml | Should Match $expected;
        }
    }
}

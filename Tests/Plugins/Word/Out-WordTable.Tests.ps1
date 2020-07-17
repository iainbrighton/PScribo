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

    Describe 'Plugins\Word\Out-WordTable' {

        BeforeEach {
            $testRows = Get-Process |
                Select-Object -Property 'ProcessName','SI','Id' -First 3
        }

        foreach ($tableStyle in @($false, $true)) {

            $tableType = if ($tableStyle) { 'Tabular' } else { 'List' }

            It "outputs $tableType table border `"<w:tblBorders>[..]</w:tblBorders>`"" {
                $document = Document -Name 'TestDocument' {
                    $testRows |
                            Table -Name 'Test Table' -List:$tableStyle
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tblBorders>[..]</w:tblBorders>'
                $testDocument.DocumentElement.OuterXml | Should Match $expected
            }

            It "outputs $tableType table border color" {
                $defaulBorderColor = '2A70BE'
                $document = Document -Name 'TestDocument' {
                    $testRows |
                            Table -Name 'Test Table' -List:$tableStyle
                }

                $testDocument = Get-WordDocument -Document $document

                $testDocument.OuterXml | Should Match (GetMatch ('<w:top w:sz="5.76" w:val="single" w:color="2A70BE" />' -f $defaulBorderColor))
                $testDocument.OuterXml | Should Match (GetMatch ('<w:bottom w:sz="5.76" w:val="single" w:color="{0}" />' -f $defaulBorderColor))
                $testDocument.OuterXml | Should Match (GetMatch ('<w:start w:sz="5.76" w:val="single" w:color="{0}" />' -f $defaulBorderColor))
                $testDocument.OuterXml | Should Match (GetMatch ('<w:end w:sz="5.76" w:val="single" w:color="{0}" />' -f $defaulBorderColor))
                $testDocument.OuterXml | Should Match (GetMatch ('<w:insideH w:sz="5.76" w:val="single" w:color="{0}" />' -f $defaulBorderColor))
                $testDocument.OuterXml | Should Match (GetMatch ('<w:insideV w:sz="5.76" w:val="single" w:color="{0}" />' -f $defaulBorderColor))
            }

            It "outputs $tableType table cell spacing `"<w:tblCellMar>[..]</w:tblCellMar>`"" {
                $defaultPaddingTop = ConvertTo-InvariantCultureString -Object 14.4
                $defaultPaddingLeft = ConvertTo-InvariantCultureString -Object 86.4
                $defaultPaddingBottom = ConvertTo-InvariantCultureString -Object 0
                $defaultPaddingRight = ConvertTo-InvariantCultureString -Object 86.4
                $document = Document -Name 'TestDocument' {
                    $testRows |
                            Table -Name 'Test Table' -List:$tableStyle
                }

                $testDocument = Get-WordDocument -Document $document

                $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:tblCellMar>[..]</w:tblCellMar>'))
                $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:top w:w="{0}" w:type="dxa" />' -f $defaultPaddingTop))
                $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:start w:w="{0}" w:type="dxa" />' -f $defaultPaddingLeft))
                $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:bottom w:w="{0}" w:type="dxa" />' -f $defaultPaddingBottom))
                $testDocument.DocumentElement.OuterXml | Should Match (GetMatch ('<w:end w:w="{0}" w:type="dxa" />' -f $defaultPaddingRight))
            }

            It "outputs $tableType table spacing `"<w:spacing w:before=`"72`" w:after=`"72`" />`"" {
                $document = Document -Name 'TestDocument' {
                    $testRows |
                            Table -Name 'Test Table' -List:$tableStyle
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('<w:spacing w:before="72" w:after="72" />')
                $testDocument.DocumentElement.OuterXml | Should Match $expected
            }

            It "outputs $tableType table with default left alignment `"<w:jc W:val=`"start`" />`"" {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:jc w:val="start" />'
                $testDocument.OuterXml | Should Match $expected
            }

            It "outputs $tableType table with center alignment `"<w:jc W:val=`"center`" />`"" {
                $document = Document -Name 'TestDocument' {
                    TableStyle 'Center' -Align Center
                    $testRows | Table -Style 'Center'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:jc w:val="center" />'
                $testDocument.OuterXml | Should Match $expected
            }

            It "outputs $tableType table with right alignment `"<w:jc W:val=`"end`" />`"" {
                $document = Document -Name 'TestDocument' {
                    TableStyle 'Right' -Align Right
                    $testRows | Table -Style 'Right'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:jc w:val="end" />'
                $testDocument.OuterXml | Should Match $expected
            }

        }

        Context 'List Table' {

            BeforeEach {
                $testRows = Get-Process |
                    Select-Object -Property 'ProcessName','SI','Id' -First 3
            }

            It 'outputs table per row "(<w:tbl>[..]</w:tbl>.*){3}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -List
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:tbl>[..]</w:tbl>.*){{{0}}}' -f ($testRows.Count))
                $testDocument.DocumentElement.OuterXml | Should Match $expected
            }

            It 'outputs space between each table "(<w:p />.*){2}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -List
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:p />.*){{{0}}}' -f ($testRows.Count -1))
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs one row per object property "(<w:tr>[..]</w:tr>.*){3}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -List
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:tr>[..]</w:tr>.*){{{0}}}' -f $testRows.Count)
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs two cells per object property "(<w:tc>[..]</w:tc>.*){6}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -List;
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:tc>[..]</w:tc>.*){{{0}}}' -f ($testRows.Count * 2))
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table cell percentage widths "(<w:tcW w:w="[..]" w:type="pct" />.*){6}' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Select-Object -First 1 |
                        Table -Name 'Test Table' -List -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:tcW w:w="[..]" w:type="pct" />.*){{{0}}}' -f ($testRows.Count * 2))
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs paragraph per table cell "(<w:p>[..]</w:p>.*){6}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Select-Object -First 1 |
                        Table -Name 'Test Table' -List -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:p>[..]</w:p>.*){6}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs custom cell style "(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>)"' {
                $testStyleName = 'Custom'
                $document = Document -Name 'TestDocument' {
                    Style -Name $testStyleName -Bold
                    $testTable = @(
                        [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                        [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; 'Column2__Style' = $testStyleName; }
                    )
                    Table 'TestTable' -Hashtable $testTable -List
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>){{1}}' -f $testStyleName)
                $testDocument.OuterXml | Should Match $expected
            }
        }

        Context 'Tabular Table' {

            BeforeEach {
                $testRows = Get-Process |
                    Select-Object -Property 'ProcessName','SI','Id' -First 3
            }

            It 'appends table "<w:tbl>[..]</w:tbl>"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tbl>[..]</w:tbl>'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table rows including header "(<w:tr>[..]?</w:tr>){4}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:tr>[..]?</w:tr>){4}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table header "<w:tr><w:trPr><w:tblHeader />"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tr><w:trPr><w:tblHeader />'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table borders "<w:tblBorders><w:top[..]/><w:bottom[..]/><w:start[..]/>[..]</w:tblBorders>"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tblBorders><w:top [..] /><w:bottom [..] /><w:start [..] />[..]</w:tblBorders>'
                $testDocument.OuterXml | Should Match $expected
            }
#
            It 'outputs table borders "<w:tblBorders>[..]<w:end[..]/><w:insideH[..]/><w:insideV[..]/></w:tblBorders>"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tblBorders>[..]<w:end [..] /><w:insideH [..] /><w:insideV [..] /></w:tblBorders>'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs grid per column "(<w:gridCol w:w=`"\d+`" />){2}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:gridCol w:w="\d+" />){2}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table cell percentage widths "(<w:tcW w:w="[..]" w:type="pct" />.*){2}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:tcW w:w="[..]" w:type="pct" />.*){2}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs cell style "(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>){1}"' {
                $testStyleName = 'Custom'
                $document = Document -Name 'TestDocument' {
                    Style -Name $testStyleName -Bold
                    $testTable = @(
                        [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                        [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; 'Column2__Style' = $testStyleName; }
                    )
                    Table 'TestTable' -Hashtable $testTable -List
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>){{1}}' -f $testStyleName)
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs row style "(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>.*){2}"' {
                $testStyleName = 'Custom'
                $document = Document -Name 'TestDocument' {
                    Style -Name $testStyleName -Bold
                    $testTable = @(
                        [Ordered] @{ Column1 = 'Row1/Column1'; Column2 = 'Row1/Column2'; }
                        [Ordered] @{ Column1 = 'Row2/Column1'; Column2 = 'Row2/Column2'; '__Style' = $testStyleName; }
                    )
                    Table 'TestTable' -Hashtable $testTable
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch ('(<w:pPr><w:pStyle w:val="{0}" /><w:spacing w:before="0" w:after="0" /></w:pPr>.*){{2}}' -f $testStyleName)
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table cells per row "(<w:tc>[..]?<\/w:tc>.*){8}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:tc>[..]?</w:tc>.*){8}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs paragraph per table cell "(<w:p>[..]</w:p>.*){8}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:p>[..]</w:p>.*){8}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs run per table cell "(<w:r>[..]</w:r>.*){8}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:r>[..]</w:r>.*){8}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs text per cell "(<w:t>[..]</w:t>.*){8}"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '(<w:t>[..]</w:t>.*){8}'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs default table style "<w:tblStyle w:val="TableDefault" />"' {
                $document = Document -Name 'TestDocument' {
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tblStyle w:val="TableDefault" />'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs custom table style "<w:tblStyle w:val="CustomTableStyle" />"' {
                $document = Document -Name 'TestDocument' {
                    Style -Name 'CustomStyle' -Size 11 -Color 000;
                    TableStyle -Id 'CustomTableStyle' -HeaderStyle 'CustomStyle' -RowStyle 'CustomStyle' -AlternateRowStyle 'CustomStyle';
                    $testRows | Table -Name 'Test Table' -Columns 'ProcessName','SI' -ColumnWidths 30,70 -Style 'CustomTableStyle'
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch '<w:tblStyle w:val="CustomTableStyle" />'
                $testDocument.OuterXml | Should Match $expected
            }

            It 'outputs table cell with embedded new line' {
                $licenses = 'Standard{0}Professional{0}Enterprise' -f [System.Environment]::NewLine
                $document = Document -Name 'TestDocument' {
                    Table -Hashtable ([Ordered] @{ Licenses = $licenses })
                }

                $testDocument = Get-WordDocument -Document $document

                $expected = GetMatch "(<w:t>[..]</w:t><w:br />){2}<w:t>[..]</w:t>"
                $testDocument.OuterXml | Should Match $expected
            }
        }
    }
}

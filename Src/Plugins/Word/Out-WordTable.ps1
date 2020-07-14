function Out-WordTable
{
<#
    .SYNOPSIS
        Output one (or more listed) formatted Word tables.

    .NOTES
        Specifies that the current row should be repeated at the top each new page on which the table is displayed. E.g, <w:tblHeader />.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table,

        ## Root element to append the table(s) to. List view will create multiple tables
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Xml.XmlElement] $Element,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    begin
    {
        $formattedTables = ConvertTo-PScriboPreformattedTable -Table $Table
    }
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style

        foreach ($formattedTable in $formattedTables)
        {
            if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Above'))
            {
                $tableCaption = (Get-WordTableCaption -Table $Table -XmlDocument $XmlDocument)
                [ref] $null = $Element.AppendChild($tableCaption)
            }

            $tbl = $Element.AppendChild($XmlDocument.CreateElement('w', 'tbl', $xmlns))
            [ref] $null = $tbl.AppendChild((Get-WordTablePr -Table $Table -XmlDocument $XmlDocument))

            $tblGrid = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblGrid', $xmlns))
            for ($i = 0; $i -lt $Table.Columns.Count; $i++)
            {
                [ref] $null = $tblGrid.AppendChild($XmlDocument.CreateElement('w', 'gridCol', $xmlns))
            }

            for ($r = 0; $r -lt $formattedTable.Rows.Count; $r++)
            {
                $row = $formattedTable.Rows[$r]
                $isRowStyleInherited = $row.IsStyleInherited
                $rowStyle = $null
                if (-not $isRowStyleInherited) {
                    $rowStyle = Get-PScriboDocumentStyle -Style $row.Style
                }

                $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlns))

                if (($r -eq 0) -and ($formattedTable.HasHeaderRow))
                {
                    $trPr = $tr.AppendChild($XmlDocument.CreateElement('w', 'trPr', $xmlns))
                    [ref] $null = $trPr.AppendChild($XmlDocument.CreateElement('w', 'tblHeader', $xmlns))
                    $cnfStyle = $trPr.AppendChild($XmlDocument.CreateElement('w', 'cnfStyle', $xmlns))
                    [ref] $null = $cnfStyle.SetAttribute('firstRow', $xmlns, 1)
                }

                for ($c = 0; $c -lt $row.Cells.Count; $c++)
                {
                    $cell = $row.Cells[$c]
                    $isCellStyleInherited = $cell.IsStyleInherited
                    $cellStyle = $null
                    if (-not $isCellStyleInherited)
                    {
                        $cellStyle = Get-PScriboDocumentStyle -Style $cell.Style
                    }

                    $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlns))
                    $tcPr = $tc.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlns))

                    if ((-not $isCellStyleInherited) -and (-not [System.String]::IsNullOrEmpty($cellStyle.BackgroundColor)))
                    {
                        $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlns))
                        [ref] $null = $shd.SetAttribute('val', $xmlns, 'clear')
                        [ref] $null = $shd.SetAttribute('color', $xmlns, 'auto')
                        $backgroundColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $cellStyle.BackgroundColor)
                        [ref] $null = $shd.SetAttribute('fill', $xmlns, $backgroundColor)
                    }
                    elseif ((-not $isRowStyleInherited) -and (-not [System.String]::IsNullOrEmpty($rowStyle.BackgroundColor)))
                    {
                        $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlns))
                        [ref] $null = $shd.SetAttribute('val', $xmlns, 'clear')
                        [ref] $null = $shd.SetAttribute('color', $xmlns, 'auto')
                        $backgroundColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $rowStyle.BackgroundColor)
                        [ref] $null = $shd.SetAttribute('fill', $xmlns, $backgroundColor)
                    }

                    if (($c -eq 0) -and ($r -ne 0))
                    {
                        $cnfStyle = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'cnfStyle', $xmlns))
                        [ref] $null = $cnfStyle.SetAttribute('firstColumn', $xmlns, 1)
                    }

                    $tcW = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlns))
                    if (($null -ne $Table.ColumnWidths) -and ($null -ne $Table.ColumnWidths[$c]))
                    {
                        [ref] $null = $tcW.SetAttribute('w', $xmlns, ($Table.ColumnWidths[$c] * 50))
                        [ref] $null = $tcW.SetAttribute('type', $xmlns, 'pct')
                    }
                    else
                    {
                        [ref] $null = $tcW.SetAttribute('w', $xmlns, 0)
                        [ref] $null = $tcW.SetAttribute('type', $xmlns, 'auto')

                    }

                    ## Scaffold paragraph and paragraph run for cell content
                    $paragraph = [PSCustomObject] @{
                        Id                = [System.Guid]::NewGuid().ToString()
                        Type              = 'PScribo.Paragraph'
                        Style             = $null
                        Tabs              = 0
                        Sections          = (New-Object -TypeName System.Collections.ArrayList)
                        IsSectionBreakEnd = $false
                    }
                    if (-not $isCellStyleInherited)
                    {
                        $paragraph.Style = $cellStyle.Id
                    }
                    elseif (-not $isRowStyleInherited)
                    {
                        $paragraph.Style = $rowStyle.Id
                    }
                    $paragraphRun = New-PScriboParagraphRun -Text $cell.Content
                    [ref] $null = $paragraph.Sections.Add($paragraphRun)
                    $p = Out-WordParagraph -Paragraph $paragraph -XmlDocument $XmlDocument
                    [ref] $null = $tc.AppendChild($p)
                }
            }

            if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Below'))
            {
                $tableCaption = Get-WordTableCaption -Table $Table -XmlDocument $XmlDocument
                [ref] $null = $Element.AppendChild($tableCaption)
            }

            ## Output empty line after (each) table
            [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        }
    }
}

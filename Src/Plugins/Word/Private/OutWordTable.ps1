function OutWordTable
{
<#
    .SYNOPSIS
        Output formatted Word table.

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
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $tableStyle = $Document.TableStyles[$Table.Style]
        $headerStyle = $Document.Styles[$tableStyle.HeaderStyle]

        if ($Table.IsList)
        {
            for ($r = 0; $r -lt $Table.Rows.Count; $r++)
            {
                $row = $Table.Rows[$r]
                if ($r -gt 0)
                {
                    ## Add a space between each table as Word renders them together..
                    [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                }

                ## Create <tr><tc></tc></tr> for each property
                $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument))

                $properties = @($row.PSObject.Properties)
                for ($i = 0; $i -lt $properties.Count; $i++)
                {
                    $propertyName = $properties[$i].Name
                    ## Ignore __Style properties
                    if (-not $propertyName.EndsWith('__Style', 'CurrentCultureIgnoreCase'))
                    {
                        $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain))
                        $tc1 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain))
                        $tcPr1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain))

                        if ($null -ne $Table.ColumnWidths)
                        {
                            ## TODO: Refactor out
                            [ref] $null = ConvertTo-Twips -Millimeter $Table.ColumnWidths[0]
                            $tcW1 = $tcPr1.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain))
                            [ref] $null = $tcW1.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[0] * 50)
                            [ref] $null = $tcW1.SetAttribute('type', $xmlnsMain, 'pct')
                        }

                        if ($headerStyle.BackgroundColor)
                        {
                            [ref] $null = $tc1.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument))
                        }

                        $p1 = $tc1.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                        $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                        $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                        [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle)
                        $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                        $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                        [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($propertyName))

                        $tc2 = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain))
                        $tcPr2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain))

                        if ($null -ne $Table.ColumnWidths)
                        {
                            ## TODO: Refactor out
                            $tcW2 = $tcPr2.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain))
                            [ref] $null = $tcW2.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[1] * 50)
                            [ref] $null = $tcW2.SetAttribute('type', $xmlnsMain, 'pct')
                        }

                        $p2 = $tc2.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                        $cellPropertyStyle = '{0}__Style' -f $propertyName
                        if ($row.PSObject.Properties[$cellPropertyStyle])
                        {
                            if (-not (Test-Path -Path Variable:\cellStyle))
                            {
                                $cellStyle = $Document.Styles[$row.$cellPropertyStyle]
                            }
                            elseif ($cellStyle.Id -ne $row.$cellPropertyStyle)
                            {
                                ## Retrieve the style if we don't already have it
                                $cellStyle = $Document.Styles[$row.$cellPropertyStyle]
                            }

                            if ($cellStyle.BackgroundColor)
                            {
                                [ref] $null = $tc2.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument))
                            }

                            if ($row.$cellPropertyStyle)
                            {
                                $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                                $pStyle2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                                [ref] $null = $pStyle2.SetAttribute('val', $xmlnsMain, $row.$cellPropertyStyle)
                            }
                        }

                        if ($null -ne $row.($propertyName))
                        {
                            ## Create a separate run for each line/break
                            $lines = $row.($propertyName).ToString() -split [System.Environment]::NewLine;
                            for ($l = 0; $l -lt $lines.Count; $l++)
                            {
                                $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t2 = $r2.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode($lines[$l]));
                                if ($l -lt ($lines.Count -1))
                                {
                                    ## Don't add a line break to the last line/break
                                    $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                    $t3 = $r3.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                    [ref] $null = $t3.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                                }
                            }
                        }
                    }
                }
            }
        } #end if Table.IsList
        else
        {
            $tbl = $Element.AppendChild((GetWordTable -Table $Table -XmlDocument $XmlDocument))

            $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain))
            $trPr = $tr.AppendChild($XmlDocument.CreateElement('w', 'trPr', $xmlnsMain))
            $null = $trPr.AppendChild($XmlDocument.CreateElement('w', 'tblHeader', $xmlnsMain))
            ## Flow headers across pages
            for ($i = 0; $i -lt $Table.Columns.Count; $i++)
            {
                $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain))
                if ($headerStyle.BackgroundColor)
                {
                    $tcPr = $tc.AppendChild((GetWordTableStyleCellPr -Style $headerStyle -XmlDocument $XmlDocument))
                }
                else
                {
                    $tcPr = $tc.AppendChild($XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain))
                }
                $tcW = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'tcW', $xmlnsMain))

                if (($null -ne $Table.ColumnWidths) -and ($null -ne $Table.ColumnWidths[$i]))
                {
                    [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, $Table.ColumnWidths[$i] * 50)
                    [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'pct')
                }
                else
                {
                    [ref] $null = $tcW.SetAttribute('w', $xmlnsMain, 0)
                    [ref] $null = $tcW.SetAttribute('type', $xmlnsMain, 'auto')
                }

                $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $tableStyle.HeaderStyle)
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($Table.Columns[$i]))
            } #end for Table.Columns

            $isAlternatingRow = $false
            foreach ($row in $Table.Rows)
            {
                $tr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tr', $xmlnsMain))
                foreach ($propertyName in $Table.Columns)
                {
                    $cellPropertyStyle = '{0}__Style' -f $propertyName
                    if ($row.PSObject.Properties[$cellPropertyStyle])
                    {
                        ## Cell style overrides row/default styles
                        $cellStyleName = $row.$cellPropertyStyle
                    }
                    elseif (-not [System.String]::IsNullOrEmpty($row.__Style))
                    {
                        ## Row style overrides default style
                        $cellStyleName = $row.__Style
                    }
                    else
                    {
                        ## Use the table row/alternating style..
                        $cellStyleName = $tableStyle.RowStyle
                        if ($isAlternatingRow) {
                            $cellStyleName = $tableStyle.AlternateRowStyle
                        }
                    }

                    if (-not (Test-Path -Path Variable:\cellStyle))
                    {
                        $cellStyle = $Document.Styles[$cellStyleName]
                    }
                    elseif ($cellStyle.Id -ne $cellStyleName)
                    {
                        ## Retrieve the style if we don't already have it
                        $cellStyle = $Document.Styles[$cellStyleName]
                    }

                    $tc = $tr.AppendChild($XmlDocument.CreateElement('w', 'tc', $xmlnsMain))
                    if ($cellStyle.BackgroundColor) {
                        [ref] $null = $tc.AppendChild((GetWordTableStyleCellPr -Style $cellStyle -XmlDocument $XmlDocument))
                    }
                    $p = $tc.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                    $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                    [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $cellStyleName)

                    if ($null -ne $row.($propertyName))
                    {
                        ## Create a separate run for each line/break
                        $lines = $row.($propertyName).ToString() -split [System.Environment]::NewLine;
                        for ($l = 0; $l -lt $lines.Count; $l++)
                        {
                            $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                            $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                            [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]));
                            if ($l -lt ($lines.Count -1))
                            {
                                ## Don't add a line break to the last line/break
                                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain));
                                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain));
                                [ref] $null = $t.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain));
                            }
                        } #end foreach line break
                    }
                } #end foreach property
                $isAlternatingRow = !$isAlternatingRow
            } #end foreach row
        } #end if not Table.IsList
    }
}

        #region OutWord Private Functions

        function ConvertToWordColor {
        <#
            .SYNOPSIS
                Converts an HTML color to RRGGBB value as Word does not support short Html color codes
        #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.String] $Color
            )
            process
            {
                $Color = $Color.TrimStart('#')
                if ($Color.Length -eq 3)
                {
                    $Color = '{0}{0}{1}{1}{2}{2}' -f $Color[0], $Color[1], $Color[2]
                }
                return $Color.ToUpper()
            }
        }


        function OutWordSection {
        <#
            .SYNOPSIS
                Output formatted Word section (paragraph).
        #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Section,

                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $RootElement,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

                $p = $RootElement.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain));
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));

                if (-not [System.String]::IsNullOrEmpty($Section.Style))
                {
                    #if (-not $Section.IsExcluded) {
                    ## If it's excluded we need a non-Heading style :( Could explicitly set the style on the run?
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                    [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Section.Style)
                    #}
                }

                if ($Section.Tabs -gt 0)
                {
                    $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain));
                    [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Section.Tabs));
                }

                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain));
                ## Increment heading spacing by 2pt for each section level, starting at 8pt for level 0, 10pt for level 1 etc
                $spacingPt = (($Section.Level * 2) + 8) * 20
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, $spacingPt)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, $spacingPt)
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))

                if ($Document.Options['EnableSectionNumbering'])
                {
                    [System.String] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name
                }
                else
                {
                    [System.String] $sectionName = '{0}' -f $Section.Name
                }

                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($sectionName))

                foreach ($s in $Section.Sections.GetEnumerator())
                {
                    if ($s.Id.Length -gt 40)
                    {
                        $sectionId = '{0}[..]' -f $s.Id.Substring(0, 36)
                    }
                    else
                    {
                        $sectionId = $s.Id
                    }
                    $currentIndentationLevel = 1
                    if ($null -ne $s.PSObject.Properties['Level'])
                    {
                        $currentIndentationLevel = $s.Level + 1
                    }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel
                    switch ($s.Type)
                    {
                        'PScribo.Section' {
                            $s | OutWordSection -RootElement $RootElement -XmlDocument $XmlDocument
                        }
                        'PScribo.Paragraph' {
                            [ref] $null = $RootElement.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $XmlDocument))
                        }
                        'PScribo.PageBreak' {
                            [ref] $null = $RootElement.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $XmlDocument))
                        }
                        'PScribo.LineBreak' {
                            [ref] $null = $RootElement.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $XmlDocument))
                        }
                        'PScribo.Table' {
                            OutWordTable -Table $s -XmlDocument $XmlDocument -Element $RootElement
                        }
                        'PScribo.BlankLine' {
                            OutWordBlankLine -BlankLine $s -XmlDocument $XmlDocument -Element $RootElement
                        }
                        'PScribo.Image' {
                            [ref] $null = $RootElement.AppendChild((OutWordImage -Image $s -XmlDocument $XmlDocument))
                        }
                        Default {
                            WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning
                        }
                    }
                }

                if ($Section.IsSectionBreakEnd)
                {
                    $sectionPrParams = @{
                        PageHeight       = if ($Section.Orientation -eq 'Portrait') { $Document.Options['PageHeight'] } else { $Document.Options['PageWidth'] }
                        PageWidth        = if ($Section.Orientation -eq 'Portrait') { $Document.Options['PageWidth'] } else { $Document.Options['PageHeight'] }
                        PageMarginTop    = $Document.Options['MarginTop'];
                        PageMarginBottom = $Document.Options['MarginBottom'];
                        PageMarginLeft   = $Document.Options['MarginLeft'];
                        PageMarginRight  = $Document.Options['MarginRight'];
                        Orientation      = $Section.Orientation;
                    }
                    [ref] $null = $pPr.AppendChild((GetWordSectionPr @sectionPrParams -XmlDocument $xmlDocument));
                }
            }
        }


        function OutWordParagraph {
        <#
            .SYNOPSIS
                Output formatted Word paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Paragraph,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain);
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain));

                if ($Paragraph.Tabs -gt 0)
                {
                    $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlnsMain))
                    [ref] $null = $ind.SetAttribute('left', $xmlnsMain, (720 * $Paragraph.Tabs))
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
                {
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                    [ref] $null = $pStyle.SetAttribute('val', $xmlnsMain, $Paragraph.Style)
                }

                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)

                if ([System.String]::IsNullOrEmpty($Paragraph.Text))
                {
                    $lines = $Paragraph.Id -Split [System.Environment]::NewLine
                }
                else
                {
                    $lines = $Paragraph.TexT -Split [System.Environment]::NewLine
                }

                ## Create a separate run for each line/break
                for ($l = 0; $l -lt $lines.Count; $l++)
                {
                    $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                    $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain))
                    ## Apply custom paragraph styles to the run..
                    if ($Paragraph.Font)
                    {
                        $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain))
                        [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Paragraph.Font[0])
                        [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Paragraph.Font[0])
                    }
                    if ($Paragraph.Size -gt 0)
                    {
                        $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain))
                        [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Paragraph.Size * 2)
                    }
                    if ($Paragraph.Bold -eq $true)
                    {
                        [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain))
                    }
                    if ($Paragraph.Italic -eq $true)
                    {
                        [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain))
                    }
                    if ($Paragraph.Underline -eq $true)
                    {
                        $u = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain))
                        [ref] $null = $u.SetAttribute('val', $xmlnsMain, 'single')
                    }
                    if (-not [System.String]::IsNullOrEmpty($Paragraph.Color))
                    {
                        $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
                        [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Paragraph.Color))
                    }

                    $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                    [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
                    ## needs to be xml:space="preserve" NOT w:space...
                    [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($lines[$l]))

                    if ($l -lt ($lines.Count - 1))
                    {
                        ## Don't add a line break to the last line/break
                        $brr = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                        $brt = $brr.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                        [ref] $null = $brt.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain))
                    }
                }

                if ($Paragraph.IsSectionBreakEnd)
                {
                    $paragraphPrParams = @{
                        PageHeight       = if ($Paragraph.Orientation -eq 'Portrait') { $Document.Options['PageHeight'] } else { $Document.Options['PageWidth'] }
                        PageWidth        = if ($Paragraph.Orientation -eq 'Portrait') { $Document.Options['PageWidth'] } else { $Document.Options['PageHeight'] }
                        PageMarginTop    = $Document.Options['MarginTop'];
                        PageMarginBottom = $Document.Options['MarginBottom'];
                        PageMarginLeft   = $Document.Options['MarginLeft'];
                        PageMarginRight  = $Document.Options['MarginRight'];
                        Orientation      = $Paragraph.Orientation;
                    }
                    [ref] $null = $pPr.AppendChild((GetWordSectionPr @paragraphPrParams -XmlDocument $xmlDocument));
                }
                return $p;
            }
        }


        function OutWordPageBreak {
        <#
            .SYNOPSIS
            Output formatted Word page break.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $PageBreak,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $br = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlnsMain))
                [ref] $null = $br.SetAttribute('type', $xmlnsMain, 'page')
                return $p
            }
        }


        function OutWordLineBreak {
        <#
            .SYNOPSIS
            Output formatted Word line break.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $LineBreak,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                $pBdr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pBdr', $xmlnsMain))
                $bottom = $pBdr.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain))
                [ref] $null = $bottom.SetAttribute('val', $xmlnsMain, 'single')
                [ref] $null = $bottom.SetAttribute('sz', $xmlnsMain, 6)
                [ref] $null = $bottom.SetAttribute('space', $xmlnsMain, 1)
                [ref] $null = $bottom.SetAttribute('color', $xmlnsMain, 'auto')
                return $p
            }
        }


        function GetWordTable {
        <#
            .SYNOPSIS
            Creates a scaffold Word <w:tbl> element
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Management.Automation.PSObject] $Table,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $tableStyle = $Document.TableStyles[$Table.Style]
                $tbl = $XmlDocument.CreateElement('w', 'tbl', $xmlnsMain)
                $tblPr = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))

                if ($Table.Tabs -gt 0)
                {
                    $tblInd = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblInd', $xmlnsMain))
                    [ref] $null = $tblInd.SetAttribute('w', $xmlnsMain, (720 * $Table.Tabs))
                }

                if ($Table.ColumnWidths)
                {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain))
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'fixed')
                }
                elseif ($Table.Width -eq 0)
                {
                    $tblLayout = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLayout', $xmlnsMain))
                    [ref] $null = $tblLayout.SetAttribute('type', $xmlnsMain, 'autofit')
                }

                if ($Table.Width -gt 0)
                {
                    $tblW = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblW', $xmlnsMain))
                    [ref] $null = $tblW.SetAttribute('type', $xmlnsMain, 'pct')
                    $tableWidthRenderPct = $Table.Width

                    if ($Table.Tabs -gt 0)
                    {
                        ## We now need to deal with tables being pushed outside the page margin
                        $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options['PageMarginLeft'] + $Document.Options['PageMarginRight'])
                        $indentWidthMm = ConvertPtToMm -Point ($Table.Tabs * 36)
                        $tableRenderMm = (($pageWidthMm / 100) * $Table.Width) + $indentWidthMm
                        if ($tableRenderMm -gt $pageWidthMm)
                        {
                            ## We've over-flowed so need to work out the maximum percentage
                            $maxTableWidthMm = $pageWidthMm - $indentWidthMm
                            $tableWidthRenderPct = [System.Math]::Round(($maxTableWidthMm / $pageWidthMm) * 100, 2)
                            WriteLog -Message ($localized.TableWidthOverflowWarning -f $tableWidthRenderPct) -IsWarning
                        }
                    }
                    [ref] $null = $tblW.SetAttribute('w', $xmlnsMain, $tableWidthRenderPct * 50)
                }

                $spacing = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 72)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 72)

                #$tblLook = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblLook', $xmlnsMain));
                #[ref] $null = $tblLook.SetAttribute('val', $xmlnsMain, '04A0');
                #[ref] $null = $tblLook.SetAttribute('firstRow', $xmlnsMain, 1);
                ## <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
                #$tblStyle = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyle', $xmlnsMain));
                #[ref] $null = $tblStyle.SetAttribute('val', $xmlnsMain, $Table.Style);

                if ($tableStyle.BorderWidth -gt 0)
                {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain))
                    foreach ($border in @('top', 'bottom', 'start', 'end', 'insideH', 'insideV'))
                    {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain))
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth))
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single')
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor))
                    }
                }

                $tblCellMar = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblCellMar', $xmlnsMain))
                $top = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'top', $xmlnsMain))
                [ref] $null = $top.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingTop))
                [ref] $null = $top.SetAttribute('type', $xmlnsMain, 'dxa')
                $left = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlnsMain))
                [ref] $null = $left.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingLeft))
                [ref] $null = $left.SetAttribute('type', $xmlnsMain, 'dxa')
                $bottom = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'bottom', $xmlnsMain))
                [ref] $null = $bottom.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingBottom))
                [ref] $null = $bottom.SetAttribute('type', $xmlnsMain, 'dxa')
                $right = $tblCellMar.AppendChild($XmlDocument.CreateElement('w', 'end', $xmlnsMain))
                [ref] $null = $right.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips $tableStyle.PaddingRight))
                [ref] $null = $right.SetAttribute('type', $xmlnsMain, 'dxa')

                $tblGrid = $tbl.AppendChild($XmlDocument.CreateElement('w', 'tblGrid', $xmlnsMain))
                for ($i = 0; $i -lt $Table.Columns.Count; $i++)
                {
                    [ref] $null = $tblGrid.AppendChild($XmlDocument.CreateElement('w', 'gridCol', $xmlnsMain))
                }

                return $tbl
            }
        }


        function OutWordTable {
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
                                    [ref] $null = ConvertMmToTwips -Millimeter $Table.ColumnWidths[0]
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



        function OutWordTOC {
        <#
            .SYNOPSIS
                Output formatted Word table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $TOC,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlnsMain)
                $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr', $xmlnsMain))
                $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', 'docPartObj', $xmlnsMain))
                $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartGallery', $xmlnsMain))
                [ref] $null = $docObjectGallery.SetAttribute('val', $xmlnsMain, 'Table of Contents')
                [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartUnique', $xmlnsMain))
                [ref] $null = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdEndPr', $xmlnsMain))

                $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdContent', $xmlnsMain))
                $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
                [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, 'TOC')
                $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
                [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name))

                $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs', $xmlnsMain))
                $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', $xmlnsMain))
                [ref] $null = $tab2.SetAttribute('val', $xmlnsMain, 'right')
                [ref] $null = $tab2.SetAttribute('leader', $xmlnsMain, 'dot')
                [ref] $null = $tab2.SetAttribute('pos', $xmlnsMain, '9016')
                #10790?!
                $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                ##TODO: Refactor duplicate code
                $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
                [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlnsMain, 'begin')

                $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlnsMain))
                [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
                [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' TOC \o "1-3" \h \z \u '))

                $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
                [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlnsMain, 'separate')

                $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                #$rPr3 = $r3.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
                $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
                [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlnsMain, 'end')

                return $sdt
            }
        }


        function OutWordBlankLine {
        <#
            .SYNOPSIS
                Output formatted Word xml blank line (paragraph).
        #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $BlankLine,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,

                [Parameter(Mandatory)]
                [System.Xml.XmlElement] $Element
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {
                    [ref] $null = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
                }
            }
        }


        function GetWordStyle {
        <#
            .SYNOPSIS
                Generates Word Xml style element from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument,

                [Parameter(Mandatory)]
                [ValidateSet('Paragraph', 'Character')]
                [System.String] $Type
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                if ($Type -eq 'Paragraph')
                {
                    $styleId = $Style.Id
                    $styleName = $Style.Name
                    $linkId = '{0}Char' -f $Style.Id
                }
                else
                {
                    $styleId = '{0}Char' -f $Style.Id
                    $styleName = '{0} Char' -f $Style.Name
                    $linkId = $Style.Id
                }

                $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
                [ref] $null = $documentStyle.SetAttribute('type', $xmlnsMain, $Type.ToLower())

                if ($Style.Id -eq $Document.DefaultStyle)
                {
                    ## Set as default style
                    [ref] $null = $documentStyle.SetAttribute('default', $xmlnsMain, 1)
                    $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'uiPriority', $xmlnsMain))
                    [ref] $null = $uiPriority.SetAttribute('val', $xmlnsMain, 1)
                }
                elseif ($Style.Hidden -eq $true)
                {
                    ## Semi hide style (headers and footers etc)
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain))
                }
                elseif (($Document.TableStyles.Values | ForEach-Object -Process {

                            $_.HeaderStyle
                            $_.RowStyle
                            $_.AlternateRowStyle
                        }) -contains $Style.Id) {
                    ## Semi hide styles behind table styles (except default style!)
                    [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain))
                }

                [ref] $null = $documentStyle.SetAttribute('styleId', $xmlnsMain, $styleId)
                $documentStyleName = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
                [ref] $null = $documentStyleName.SetAttribute('val', $xmlnsMain, $styleName)
                $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlnsMain))
                [ref] $null = $basedOn.SetAttribute('val', $xmlnsMain, 'Normal')
                $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'link', $xmlnsMain))
                [ref] $null = $link.SetAttribute('val', $xmlnsMain, $linkId)
                $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'next', $xmlnsMain))
                [ref] $null = $next.SetAttribute('val', $xmlnsMain, 'Normal')
                [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain))
                $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain))
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain))
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
                ## Set the <w:jc> (justification) element
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))

                if ($Style.Align.ToLower() -eq 'justify')
                {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
                }
                else
                {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower())
                }

                if ($Style.BackgroundColor)
                {
                    $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain))
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear')
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto')
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor))
                }
                [ref] $null = $documentStyle.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument))

                return $documentStyle
            }
        }


        function GetWordTableStyle {
        <#
            .SYNOPSIS
                Generates Word Xml table style element from a PScribo document table style.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $tableStyle,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $Style = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
                [ref] $null = $Style.SetAttribute('type', $xmlnsMain, 'table')
                [ref] $null = $Style.SetAttribute('styleId', $xmlnsMain, $tableStyle.Id)
                $name = $Style.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
                [ref] $null = $name.SetAttribute('val', $xmlnsMain, $tableStyle.Id)
                $tblPr = $Style.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))
                $tblStyleRowBandSize = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblStyleRowBandSize', $xmlnsMain))
                [ref] $null = $tblStyleRowBandSize.SetAttribute('val', $xmlnsMain, 1)

                if ($tableStyle.BorderWidth -gt 0)
                {
                    $tblBorders = $tblPr.AppendChild($XmlDocument.CreateElement('w', 'tblBorders', $xmlnsMain))
                    foreach ($border in @('top', 'bottom', 'start', 'end', 'insideH', 'insideV'))
                    {
                        $b = $tblBorders.AppendChild($XmlDocument.CreateElement('w', $border, $xmlnsMain))
                        [ref] $null = $b.SetAttribute('sz', $xmlnsMain, (ConvertMmToOctips $tableStyle.BorderWidth))
                        [ref] $null = $b.SetAttribute('val', $xmlnsMain, 'single')
                        [ref] $null = $b.SetAttribute('color', $xmlnsMain, (ConvertToWordColor -Color $tableStyle.BorderColor))
                    }
                }

                [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.HeaderStyle] -Type Header -XmlDocument $XmlDocument))
                [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.RowStyle] -Type Row -XmlDocument $XmlDocument))
                [ref] $null = $Style.AppendChild((GetWordTableStylePr -Style $Document.Styles[$tableStyle.AlternateRowStyle] -Type AlternateRow -XmlDocument $XmlDocument))
                return $Style
            }
        }


        function GetWordStyleParagraphPr {
        <#
            .SYNOPSIS
                Generates Word paragraph (pPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $pPr = $XmlDocument.CreateElement('w', 'pPr', $xmlnsMain)
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain))
                [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain))
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))

                if ($Style.Align.ToLower() -eq 'justify')
                {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
                }
                else
                {
                    [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower())
                }
                return $pPr
            }
        }


        function GetWordStyleRunPrColor {
        <#
            .SYNOPSIS
                Generates Word run (rPr) text colour formatting property only.

            .NOTES
                This is only required to override the text colour in table rows/headers
                as I can't get this (yet) applied via the table style?
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain)
                $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
                [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color))
                return $rPr
            }
        }


        function GetWordStyleRunPr {
        <#
            .SYNOPSIS
                Generates Word run (rPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $rPr = $XmlDocument.CreateElement('w', 'rPr', $xmlnsMain)
                $rFonts = $rPr.AppendChild($XmlDocument.CreateElement('w', 'rFonts', $xmlnsMain))
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlnsMain, $Style.Font[0])
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlnsMain, $Style.Font[0])

                if ($Style.Bold)
                {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'b', $xmlnsMain))
                }
                if ($Style.Underline)
                {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'u', $xmlnsMain))
                }
                if ($Style.Italic)
                {
                    [ref] $null = $rPr.AppendChild($XmlDocument.CreateElement('w', 'i', $xmlnsMain))
                }

                $Color = $rPr.AppendChild($XmlDocument.CreateElement('w', 'color', $xmlnsMain))
                [ref] $null = $Color.SetAttribute('val', $xmlnsMain, (ConvertToWordColor -Color $Style.Color))
                $sz = $rPr.AppendChild($XmlDocument.CreateElement('w', 'sz', $xmlnsMain))
                [ref] $null = $sz.SetAttribute('val', $xmlnsMain, $Style.Size * 2)
                return $rPr
            }
        }


        function GetWordTableStyleCellPr {
        <#
            .SYNOPSIS
                Generates Word table cell (tcPr) formatting properties
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $tcPr = $XmlDocument.CreateElement('w', 'tcPr', $xmlnsMain)

                if ($Style.BackgroundColor)
                {
                    $shd = $tcPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain))
                    [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear')
                    [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto')
                    [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor))
                }
                return $tcPr
            }
        }


        function GetWordTableStylePr {
        <#
            .SYNOPSIS
                Generates Word table style (tblStylePr) formatting properties for specified table style type
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Management.Automation.PSObject] $Style,

                [Parameter(Mandatory)]
                [ValidateSet('Header', 'Row', 'AlternateRow')]
                [System.String] $Type,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $tblStylePr = $XmlDocument.CreateElement('w', 'tblStylePr', $xmlnsMain)
                [ref] $null = $tblStylePr.AppendChild($XmlDocument.CreateElement('w', 'tblPr', $xmlnsMain))

                switch ($Type)
                {
                    'Header' {
                        $tblStylePrType = 'firstRow'
                    }
                    'Row' {
                        $tblStylePrType = 'band2Horz'
                    }
                    'AlternateRow' {
                        $tblStylePrType = 'band1Horz'
                    }
                }

                [ref] $null = $tblStylePr.SetAttribute('type', $xmlnsMain, $tblStylePrType)
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleParagraphPr -Style $Style -XmlDocument $XmlDocument))
                [ref] $null = $tblStylePr.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument))
                [ref] $null = $tblStylePr.AppendChild((GetWordTableStyleCellPr -Style $Style -XmlDocument $XmlDocument))
                return $tblStylePr
            }
        }

        function GetWordSectionPr {
        <#
            .SYNOPSIS
                Outputs Office Open XML section element to set page size and margins.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory)]
                [System.Single] $PageWidth,

                [Parameter(Mandatory)]
                [System.Single] $PageHeight,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginTop,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginLeft,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginBottom,

                [Parameter(Mandatory)]
                [System.Single] $PageMarginRight,

                [Parameter(Mandatory)]
                [System.String] $Orientation,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
                $sectPr = $XmlDocument.CreateElement('w', 'sectPr', $xmlnsMain);
                $pgSz = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgSz', $xmlnsMain));
                [ref] $null = $pgSz.SetAttribute('w', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageWidth));
                [ref] $null = $pgSz.SetAttribute('h', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageHeight));
                [ref] $null = $pgSz.SetAttribute('orient', $xmlnsMain, $Orientation.ToLower());
                $pgMar = $sectPr.AppendChild($XmlDocument.CreateElement('w', 'pgMar', $xmlnsMain));
                [ref] $null = $pgMar.SetAttribute('top', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginTop));
                [ref] $null = $pgMar.SetAttribute('bottom', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginBottom));
                [ref] $null = $pgMar.SetAttribute('left', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginLeft));
                [ref] $null = $pgMar.SetAttribute('right', $xmlnsMain, (ConvertMmToTwips -Millimeter $PageMarginRight));
                return $sectPr;
            }
        }


        function OutWordStylesDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML style document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param
            (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $Styles,

                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Collections.Hashtable] $TableStyles
            )
            process
            {
                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $XmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
                [ref] $null = $XmlDocument.AppendChild($XmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
                $documentStyles = $XmlDocument.AppendChild($XmlDocument.CreateElement('w', 'styles', $xmlnsMain))

                ## Create default style
                $defaultStyle = $documentStyles.AppendChild($XmlDocument.CreateElement('w', 'style', $xmlnsMain))
                [ref] $null = $defaultStyle.SetAttribute('type', $xmlnsMain, 'paragraph')
                [ref] $null = $defaultStyle.SetAttribute('default', $xmlnsMain, '1')
                [ref] $null = $defaultStyle.SetAttribute('styleId', $xmlnsMain, 'Normal')
                $defaultStyleName = $defaultStyle.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
                [ref] $null = $defaultStyleName.SetAttribute('val', $xmlnsMain, 'Normal')
                [ref] $null = $defaultStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain))

                foreach ($Style in $Styles.Values)
                {
                    $documentParagraphStyle = GetWordStyle -Style $Style -XmlDocument $XmlDocument -Type Paragraph
                    [ref] $null = $documentStyles.AppendChild($documentParagraphStyle)
                    $documentCharacterStyle = GetWordStyle -Style $Style -XmlDocument $XmlDocument -Type Character
                    [ref] $null = $documentStyles.AppendChild($documentCharacterStyle)
                }

                foreach ($tableStyle in $TableStyles.Values)
                {
                    $documentTableStyle = GetWordTableStyle -TableStyle $tableStyle -XmlDocument $XmlDocument
                    [ref] $null = $documentStyles.AppendChild($documentTableStyle)
                }

                return $XmlDocument
            }
        }


        function OutWordSettingsDocument {
        <#
            .SYNOPSIS
                Outputs Office Open XML settings document
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlDocument])]
            param
            (
                [Parameter()]
                [System.Management.Automation.SwitchParameter] $UpdateFields
            )
            process
            {
                ## Create the Style.xml document
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                # <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                # xmlns:o="urn:schemas-microsoft-com:office:office"
                # xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                # xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                # xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"
                # xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                # xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                # xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                # xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
                # mc:Ignorable="w14 w15">
                $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument'
                [ref] $null = $settingsDocument.AppendChild($settingsDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
                $settings = $settingsDocument.AppendChild($settingsDocument.CreateElement('w', 'settings', $xmlnsMain))
                ## Set compatibility mode to Word 2013
                $compat = $settings.AppendChild($settingsDocument.CreateElement('w', 'compat', $xmlnsMain))
                $compatSetting = $compat.AppendChild($settingsDocument.CreateElement('w', 'compatSetting', $xmlnsMain))
                [ref] $null = $compatSetting.SetAttribute('name', $xmlnsMain, 'compatibilityMode')
                [ref] $null = $compatSetting.SetAttribute('uri', $xmlnsMain, 'http://schemas.microsoft.com/office/word')
                [ref] $null = $compatSetting.SetAttribute('val', $xmlnsMain, 15)

                if ($UpdateFields)
                {
                    $wupdateFields = $settings.AppendChild($settingsDocument.CreateElement('w', 'updateFields', $xmlnsMain))
                    [ref] $null = $wupdateFields.SetAttribute('val', $xmlnsMain, 'true')
                }
                return $settingsDocument
            }
        }


        function OutWordImage {
        <#
            .SYNOPSIS
                Output Image to Word.
        #>
            [CmdletBinding()]
            [OutputType([System.Xml.XmlElement])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Image,

                [Parameter(Mandatory)]
                [System.Xml.XmlDocument] $XmlDocument
            )
            process
            {
                $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                $xmlnswpDrawingWordProcessing = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                $xmlnsDrawingMain = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                $xmlnsDrawingPicture = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
                $xmlnsRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
                $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
                [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
                [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
                $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))
                [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Image.Align.ToLower())
                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
                $rPr = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain))
                $drawing = $rPr.AppendChild($XmlDocument.CreateElement('w', 'drawing', $xmlnsMain))
                $inline = $drawing.AppendChild($XmlDocument.CreateElement('wp', 'inline', $xmlnswpDrawingWordProcessing))
                [ref] $null = $inline.SetAttribute('distT', '0')
                [ref] $null = $inline.SetAttribute('distB', '0')
                [ref] $null = $inline.SetAttribute('distL', '0')
                [ref] $null = $inline.SetAttribute('distR', '0')

                $extent = $inline.AppendChild($XmlDocument.CreateElement('wp', 'extent', $xmlnswpDrawingWordProcessing))
                [ref] $null = $extent.SetAttribute('cx', $Image.WidthEm)
                [ref] $null = $extent.SetAttribute('cy', $Image.HeightEm)

                $effectExtent = $inline.AppendChild($XmlDocument.CreateElement('wp', 'effectExtent', $xmlnswpDrawingWordProcessing))
                [ref] $null = $effectExtent.SetAttribute('l', '0')
                [ref] $null = $effectExtent.SetAttribute('t', '0')
                [ref] $null = $effectExtent.SetAttribute('r', '0')
                [ref] $null = $effectExtent.SetAttribute('b', '0')

                $docPr = $inline.AppendChild($XmlDocument.CreateElement('wp', 'docPr', $xmlnswpDrawingWordProcessing))
                [ref] $null = $docPr.SetAttribute('id', $Image.ImageNumber)
                [ref] $null = $docPr.SetAttribute('name', $Image.Name)
                [ref] $null = $docPr.SetAttribute('descr', $Image.Name)

                $cNvGraphicFramePr = $inline.AppendChild($XmlDocument.CreateElement('wp', 'cNvGraphicFramePr', $xmlnswpDrawingWordProcessing))
                $graphicFrameLocks = $cNvGraphicFramePr.AppendChild($XmlDocument.CreateElement('a', 'graphicFrameLocks', $xmlnsDrawingMain))
                [ref] $null = $graphicFrameLocks.SetAttribute('noChangeAspect', '1')

                $graphic = $inline.AppendChild($XmlDocument.CreateElement('a', 'graphic', $xmlnsDrawingMain))
                $graphicData = $graphic.AppendChild($XmlDocument.CreateElement('a', 'graphicData', $xmlnsDrawingMain))
                [ref] $null = $graphicData.SetAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

                $pic = $graphicData.AppendChild($XmlDocument.CreateElement('pic', 'pic', $xmlnsDrawingPicture))
                $nvPicPr = $pic.AppendChild($XmlDocument.CreateElement('pic', 'nvPicPr', $xmlnsDrawingPicture))
                $cNvPr = $nvPicPr.AppendChild($XmlDocument.CreateElement('pic', 'cNvPr', $xmlnsDrawingPicture))
                [ref] $null = $cNvPr.SetAttribute('id', $Image.ImageNumber)
                [ref] $null = $cNvPr.SetAttribute('name', $Image.Name)
                [ref] $null = $cNvPr.SetAttribute('descr', $Image.Name)
                $cNvPicPr = $nvPicPr.AppendChild($XmlDocument.CreateElement('pic', 'cNvPicPr', $xmlnsDrawingPicture))
                $picLocks = $cNvPicPr.AppendChild($XmlDocument.CreateElement('a', 'picLocks', $xmlnsDrawingMain))
                [ref] $null = $picLocks.SetAttribute('noChangeAspect', '1')
                [ref] $null = $picLocks.SetAttribute('noChangeArrowheads', '1')

                $blipFill = $pic.AppendChild($XmlDocument.CreateElement('pic', 'blipFill', $xmlnsDrawingPicture))
                $blip = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'blip', $xmlnsDrawingMain))
                [ref] $null = $blip.SetAttribute('embed', $xmlnsRelationships, $Image.Name)
                [ref] $null = $blip.SetAttribute('cstate', 'print')
                $extlst = $blip.AppendChild($XmlDocument.CreateElement('a', 'extlst', $xmlnsDrawingMain))
                $ext = $extlst.AppendChild($XmlDocument.CreateElement('a', 'ext', $xmlnsDrawingMain))
                [ref] $null = $ext.SetAttribute('uri', '')
                [ref] $null = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'srcRect', $xmlnsDrawingMain))
                $stretch = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'stretch', $xmlnsDrawingMain))
                [ref] $null = $stretch.AppendChild($XmlDocument.CreateElement('a', 'fillRect', $xmlnsDrawingMain))

                $spPr = $pic.AppendChild($XmlDocument.CreateElement('pic', 'spPr', $xmlnsDrawingPicture))
                [ref] $null = $spPr.SetAttribute('bwMode', 'auto')
                $xfrm = $spPr.AppendChild($XmlDocument.CreateElement('a', 'xfrm', $xmlnsDrawingMain))
                $off = $xfrm.AppendChild($XmlDocument.CreateElement('a', 'off', $xmlnsDrawingMain))
                [ref] $null = $off.SetAttribute('x', '0')
                [ref] $null = $off.SetAttribute('y', '0')
                $ext = $xfrm.AppendChild($XmlDocument.CreateElement('a', 'ext', $xmlnsDrawingMain))
                [ref] $null = $ext.SetAttribute('cx', $Image.WidthEm)
                [ref] $null = $ext.SetAttribute('cy', $Image.HeightEm)

                $prstGeom = $spPr.AppendChild($XmlDocument.CreateElement('a', 'prstGeom', $xmlnsDrawingMain))
                [ref] $null = $prstGeom.SetAttribute('prst', 'rect')
                [ref] $null = $prstGeom.AppendChild($XmlDocument.CreateElement('a', 'avLst', $xmlnsDrawingMain))

                $null = $spPr.AppendChild($XmlDocument.CreateElement('a', 'noFill', $xmlnsDrawingMain))

                $ln = $spPr.AppendChild($XmlDocument.CreateElement('a', 'ln', $xmlnsDrawingMain))
                $null = $ln.AppendChild($XmlDocument.CreateElement('a', 'noFill', $xmlnsDrawingMain))

                return $p
            }
        }


        #endregion OutWord Private Functions

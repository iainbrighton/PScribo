        #region OutHtml Private Functions

        function New-PScriboHtmlOption {
        <#
            .SYNOPSIS
                Sets the text plugin specific formatting/output options.
            .NOTES
                All plugin options should be prefixed with the plugin name.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Collections.Hashtable])]
            param (
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Boolean] $NoPageLayoutStyle = $false
            )
            process {

                return @{
                    NoPageLayoutStyle = $NoPageLayoutStyle;
                }

            } #end process
        } #end function New-PScriboHtmlOption

        function GetHtmlStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document style
                [Parameter(Mandatory, ValueFromPipeline)] [System.Object] $Style
            )
            process {

                $styleBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $styleBuilder.AppendFormat(" font-family: '{0}';", $Style.Font -join "','");
                ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                $invariantFontSize = ConvertToInvariantCultureString -Object ($Style.Size / 12) -Format 'f2';
                [ref] $null = $styleBuilder.AppendFormat(' font-size: {0}em;', $invariantFontSize);
                [ref] $null = $styleBuilder.AppendFormat(' text-align: {0};', $Style.Align.ToLower());
                if ($Style.Bold) {

                    [ref] $null = $styleBuilder.Append(' font-weight: bold;');
                }
                else {

                    [ref] $null = $styleBuilder.Append(' font-weight: normal;');
                }
                if ($Style.Italic) {

                    [ref] $null = $styleBuilder.Append(' font-style: italic;');
                }
                if ($Style.Underline) {

                    [ref] $null = $styleBuilder.Append(' text-decoration: underline;');
                }
                if ($Style.Color.StartsWith('#')) {

                    [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Color.ToLower());
                }
                else {

                    [ref] $null = $styleBuilder.AppendFormat(' color: #{0};', $Style.Color);
                }
                if ($Style.BackgroundColor) {

                    if ($Style.BackgroundColor.StartsWith('#')) {

                        [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', $Style.BackgroundColor.ToLower());
                    }
                    else {

                        [ref] $null = $styleBuilder.AppendFormat(' background-color: #{0};', $Style.BackgroundColor.ToLower());
                    }
                }
                return $styleBuilder.ToString();

            }
        } #end function GetHtmlStyle


        function GetHtmlTableStyle {
        <#
            .SYNOPSIS
                Generates html stylesheet style attributes from a PScribo document table style.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document table style
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TableStyle
            )
            process {

                $tableStyleBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tableStyleBuilder.AppendFormat(' padding: {0}em {1}em {2}em {3}em;',
                    (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingTop)),
                        (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingRight)),
                            (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingBottom)),
                                (ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.PaddingLeft))),
                [ref] $null = $tableStyleBuilder.AppendFormat(' border-style: {0};', $TableStyle.BorderStyle.ToLower());

                if ($TableStyle.BorderWidth -gt 0) {

                    $invariantBorderWidth = ConvertToInvariantCultureString -Object (ConvertMmToEm $TableStyle.BorderWidth);
                    [ref] $null = $tableStyleBuilder.AppendFormat(' border-width: {0}em;', $invariantBorderWidth);
                    if ($TableStyle.BorderColor.Contains('#')) {

                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: {0};', $TableStyle.BorderColor);
                    }
                    else {

                        [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: #{0};', $TableStyle.BorderColor);
                    }
                }

                [ref] $null = $tableStyleBuilder.Append(' border-collapse: collapse;');
                ## <table align="center"> is deprecated in Html5
                if ($TableStyle.Align -eq 'Center') {

                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: auto;');
                }
                elseif ($TableStyle.Align -eq 'Right') {

                    [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: 0;');
                }
                return $tableStyleBuilder.ToString();

            }
        } #end function Outhtmltablestyle


        function GetHtmlTableDiv {
        <#
            .SYNOPSIS
                Generates Html <div style=..><table style=..> tags based upon table width, columns and indentation
            .NOTES
                A <div> is required to ensure that the table stays within the "page" boundaries/margins.
        #>
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $divBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.Tabs -gt 0) {

                    $invariantMarginLeft = ConvertToInvariantCultureString -Object (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs));
                    [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}em;">' -f $invariantMarginLeft);
                }
                else {

                    [ref] $null = $divBuilder.Append('<div>' -f (ConvertMmToEm -Millimeter (12.7 * $Table.Tabs)));
                }
                if ($Table.List) {

                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}-list"', $Table.Style.ToLower());
                }
                else {

                    [ref] $null = $divBuilder.AppendFormat('<table class="{0}"', $Table.Style.ToLower());
                }
                $styleElements = @();
                if ($Table.Width -gt 0) {

                    $styleElements += 'width:{0}%;' -f $Table.Width;
                }
                if ($Table.ColumnWidths) {

                    $styleElements += 'table-layout: fixed;';
                    $styleElements += 'word-break: break-word;'
                }
                if ($styleElements.Count -gt 0) {

                    [ref] $null = $divBuilder.AppendFormat(' style="{0}">', [String]::Join(' ', $styleElements));
                }
                else {

                    [ref] $null = $divBuilder.Append('>');
                }
                return $divBuilder.ToString();

            }
        } #end function GetHtmlTableDiv


        function GetHtmlTableColGroup {
        <#
            .SYNOPSIS
                Generates Html <colgroup> tags based on table column widths
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $colGroupBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.ColumnWidths) {

                    [ref] $null = $colGroupBuilder.Append('<colgroup>');
                    foreach ($columnWidth in $Table.ColumnWidths) {

                        if ($null -eq $columnWidth) {

                            [ref] $null = $colGroupBuilder.Append('<col />');
                        }
                        else {

                            [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{0}%; min-width:{0}%; width:{0}%" />', $columnWidth);
                        }
                    }

                    [ref] $null = $colGroupBuilder.AppendLine('</colgroup>');
                }
                return $colGroupBuilder.ToString();

            }
        } #end function GetHtmlTableDiv


        function OutHtmlTOC {
        <#
            .SYNOPSIS
                Generates Html table of contents.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            process {

                $tocBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $tocBuilder.AppendFormat('<h1 class="{0}">{1}</h1>', $TOC.ClassId, $TOC.Name);
                #[ref] $null = $tocBuilder.AppendLine('<table style="width: 100%;">');
                [ref] $null = $tocBuilder.AppendLine('<table>');
                foreach ($tocEntry in $Document.TOC) {

                    $sectionNumberIndent = '&nbsp;&nbsp;&nbsp;' * $tocEntry.Level;
                    if ($Document.Options['EnableSectionNumbering']) {

                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a href="#{2}" style="text-decoration: none;">{3}</a></td></tr>', $tocEntry.Number, $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                    else {

                        [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}<a href="#{1}" style="text-decoration: none;">{2}</a></td></tr>', $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine();
                    }
                }
                [ref] $null = $tocBuilder.AppendLine('</table>');
                return $tocBuilder.ToString();

            } #end process
        } #end function OutHtmlTOC


        function OutHtmlBlankLine {
        <#
            .SYNOPSIS
                Outputs html PScribo.Blankline.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $BlankLine
            )
            process {

                $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder;
                for ($i = 0; $i -lt $BlankLine.LineCount; $i++) {

                    [ref] $null = $blankLineBuilder.Append('<br />');
                }
                return $blankLineBuilder.ToString();

            } #end process
        } #end function OutHtmlBlankLine


        function OutHtmlStyle {
        <#
            .SYNOPSIS
                Generates an in-line HTML CSS stylesheet from a PScribo document styles and table styles.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## PScribo document styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $Styles,

                ## PScribo document tables styles
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [System.Collections.Hashtable] $TableStyles,

                ## Suppress page layout styling
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $NoPageLayoutStyle
            )
            process {

                $stylesBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                [ref] $null = $stylesBuilder.AppendLine('<style type="text/css">');
                if (-not $NoPageLayoutStyle) {

                    ## Add HTML page layout styling options, e.g. when emailing HTML documents
                    [ref] $null = $stylesBuilder.AppendLine('html { height: 100%; -webkit-background-size: cover; -moz-background-size: cover; -o-background-size: cover; background-size: cover; background: #f8f8f8; }');
                    [ref] $null = $stylesBuilder.Append("page { background: white; width: $($Document.Options['PageWidth'])mm; display: block; margin-top: 1em; margin-left: auto; margin-right: auto; margin-bottom: 1em; ");
                    [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }');
                    [ref] $null = $stylesBuilder.AppendLine('@media print { body, page { margin: 0; box-shadow: 0; } }');
                    [ref] $null = $stylesBuilder.AppendLine('hr { margin-top: 1.0em; }');
                }
                foreach ($style in $Styles.Keys) {

                    ## Build style
                    $htmlStyle = GetHtmlStyle -Style $Styles[$style];
                    [ref] $null = $stylesBuilder.AppendFormat(' .{0} {{{1} }}', $Styles[$style].Id, $htmlStyle).AppendLine();
                }
                foreach ($tableStyle in $TableStyles.Keys) {

                    $tStyle = $TableStyles[$tableStyle];
                    $tableStyleId = $tStyle.Id.ToLower();
                    $htmlTableStyle = GetHtmlTableStyle -TableStyle $tStyle;
                    $htmlHeaderStyle = GetHtmlStyle -Style $Styles[$tStyle.HeaderStyle];
                    $htmlRowStyle = GetHtmlStyle -Style $Styles[$tStyle.RowStyle];
                    $htmlAlternateRowStyle = GetHtmlStyle -Style $Styles[$tStyle.AlternateRowStyle];
                    ## Generate Standard table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} th {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(odd) td {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(even) td {{{1}{2} }}', $tableStyleId, $htmlAlternateRowStyle, $htmlTableStyle).AppendLine();
                    ## Generate List table styles
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(1) {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine();
                    [ref] $null = $stylesBuilder.AppendFormat(' table.{0}-list td:nth-child(2) {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine();
                } #end foreach style

                [ref] $null = $stylesBuilder.AppendLine('</style>');
                return $stylesBuilder.ToString().TrimEnd();

            } #end process
        } #end function OutHtmlStyle


        function OutHtmlSection {
        <#
            .SYNOPSIS
                Output formatted Html section.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            process {

                [System.Text.StringBuilder] $sectionBuilder = New-Object System.Text.StringBuilder;
                $encodedSectionName = [System.Net.WebUtility]::HtmlEncode($Section.Name);
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectionName; }
                else { [string] $sectionName = '{0}' -f $encodedSectionName; }
                [int] $headerLevel = $Section.Number.Split('.').Count;

                ## Html <h5> is the maximum supported level
                if ($headerLevel -gt 6) {

                    WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning;
                    $headerLevel = 6;
                }

                if ([string]::IsNullOrEmpty($Section.Style)) {

                    $className = $Document.DefaultStyle;
                }
                else {

                    $className = $Section.Style;
                }

                [ref] $null = $sectionBuilder.AppendFormat('<a name="{0}"><h{1} class="{2}">{3}</h{1}></a>', $Section.Id, $headerLevel, $className, $sectionName.TrimStart());
                foreach ($s in $Section.Sections.GetEnumerator()) {

                    if ($s.Id.Length -gt 40) {

                        $sectionId = '{0}[..]' -f $s.Id.Substring(0,36);
                    }
                    else {

                        $sectionId = $s.Id;
                    }

                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) {

                        $currentIndentationLevel = $s.Level +1;
                    }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlLineBreak)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlPageBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.Append((OutHtmlTable -Table $s)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();

            } #end process
        } # end function OutHtmlSection


        function GetHtmlParagraphStyle {
        <#
            .SYNOPSIS
                Generates html style attribute from PScribo paragraph style overrides.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Paragraph
            )
            process {

                $paragraphStyleBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Paragraph.Tabs -gt 0) {

                    ## Default to 1/2in tab spacing
                    $tabEm = ConvertToInvariantCultureString -Object (ConvertMmToEm -Millimeter (12.7 * $Paragraph.Tabs)) -Format 'f2';
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}em;', $tabEm);
                }
                if ($Paragraph.Font) {

                    [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0}';", $Paragraph.Font -Join "','");
                }
                if ($Paragraph.Size -gt 0) {

                    ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
                    $invariantParagraphSize = ConvertToInvariantCultureString -Object ($Paragraph.Size / 12) -Format 'f2';
                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' font-size: {0}em;', $invariantParagraphSize);
                }
                if ($Paragraph.Bold -eq $true) {

                    [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;');
                }
                if ($Paragraph.Italic -eq $true) {

                    [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;');
                }
                if ($Paragraph.Underline -eq $true) {

                    [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: underline;');
                }
                if (-not [System.String]::IsNullOrEmpty($Paragraph.Color) -and $Paragraph.Color.StartsWith('#')) {

                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $Paragraph.Color.ToLower());
                }
                elseif (-not [System.String]::IsNullOrEmpty($Paragraph.Color)) {

                    [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $Paragraph.Color.ToLower());
                }
                return $paragraphStyleBuilder.ToString().TrimStart();

            } #end process
        } #end function GetHtmlParagraphStyle


        function OutHtmlParagraph {
        <#
            .SYNOPSIS
                Output formatted Html paragraph.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            process {

                [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Text);
                if ([System.String]::IsNullOrEmpty($encodedText)) {

                    $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Id);
                }
                # $encodedText = $encodedText -replace [System.Environment]::NewLine, '<br />';
                $encodedText = $encodedText.Replace([System.Environment]::NewLine, '<br />');
                $customStyle = GetHtmlParagraphStyle -Paragraph $Paragraph;
                if ([System.String]::IsNullOrEmpty($Paragraph.Style) -and [System.String]::IsNullOrEmpty($customStyle)) {

                    [ref] $null = $paragraphBuilder.AppendFormat('<div>{0}</div>', $encodedText);
                }
                elseif ([System.String]::IsNullOrEmpty($customStyle)) {

                    [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">{1}</div>', $Paragraph.Style, $encodedText);
                }
                else {

                    [ref] $null = $paragraphBuilder.AppendFormat('<div style="{1}">{2}</div>', $Paragraph.Style, $customStyle, $encodedText);
                }
                return $paragraphBuilder.ToString();

            } #end process
        } #end OutHtmlParagraph


        function GetHtmlTableList {
        <#
            .SYNOPSIS
                Generates list html <table> from a PScribo.Table row object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table,

                [Parameter(Mandatory)]
                [System.Object] $Row
            )
            process {

                $listTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $listTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $listTableBuilder.Append((GetHtmlTableColGroup -Table $Table));
                [ref] $null = $listTableBuilder.Append('<tbody>');

                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {

                    $propertyName = $Table.Columns[$i];
                    [ref] $null = $listTableBuilder.AppendFormat('<tr><td>{0}</td>', $propertyName);
                    $propertyStyle = '{0}__Style' -f $propertyName;

                    if ($row.PSObject.Properties[$propertyStyle]) {

                        $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle]);
                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {

                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;</td></tr>', $propertyStyleHtml);
                        }
                        else {

                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                            $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />');
                            [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</td></tr>', $propertyStyleHtml, $encodedHtmlContent);
                        }
                    }
                    else {

                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {

                            [ref] $null = $listTableBuilder.Append('<td>&nbsp;</td></tr>');
                        }
                        else {

                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                            $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')
                            [ref] $null = $listTableBuilder.AppendFormat('<td>{0}</td></tr>', $encodedHtmlContent);
                        }
                    }
                } #end for each property
                [ref] $null = $listTableBuilder.AppendLine('</tbody></table></div>');
                return $listTableBuilder.ToString();

            } #end process
        } #end function GetHtmlTableList


        function GetHtmlTable {
        <#
            .SYNOPSIS
                Generates html <table> from a PScribo.Table object.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $standardTableBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableDiv -Table $Table));
                [ref] $null = $standardTableBuilder.Append((GetHtmlTableColGroup -Table $Table));

                ## Table headers
                [ref] $null = $standardTableBuilder.Append('<thead><tr>');
                for ($i = 0; $i -lt $Table.Columns.Count; $i++) {

                    [ref] $null = $standardTableBuilder.AppendFormat('<th>{0}</th>', $Table.Columns[$i]);
                }
                [ref] $null = $standardTableBuilder.Append('</tr></thead>');

                ## Table body
                [ref] $null = $standardTableBuilder.AppendLine('<tbody>');
                foreach ($row in $Table.Rows) {

                    [ref] $null = $standardTableBuilder.Append('<tr>');
                    foreach ($propertyName in $Table.Columns) {

                        $propertyStyle = '{0}__Style' -f $propertyName;

                        if ([string]::IsNullOrEmpty($Row.$propertyName)) {

                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode('&nbsp;');
                        }
                        else {

                            $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString());
                        }
                        $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />');

                        if ($row.PSObject.Properties[$propertyStyle]) {

                            ## Cell styles override row styles
                            $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent);
                        }
                        elseif (($row.PSObject.Properties['__Style']) -and (-not [System.String]::IsNullOrEmpty($row.__Style))) {

                            ## We have a row style
                            $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style]).Trim();
                            [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $rowStyleHtml, $encodedHtmlContent);
                        }
                        else {

                            if ($null -ne $row.$propertyName) {

                                ## Check that the property has a value
                                [ref] $null = $standardTableBuilder.AppendFormat('<td>{0}</td>', $encodedHtmlContent);
                            }
                            else {

                                [ref] $null = $standardTableBuilder.Append('<td>&nbsp</td>');
                            }
                        } #end if $row.PropertyStyle
                    } #end foreach property
                    [ref] $null = $standardTableBuilder.AppendLine('</tr>');
                } #end foreach row
                [ref] $null = $standardTableBuilder.AppendLine('</tbody></table></div>');
                return $standardTableBuilder.ToString();

            } #end process
        } #end function GetHtmlTableList


        function OutHtmlTable {
        <#
            .SYNOPSIS
                Output formatted Html <table> from PScribo.Table object.
            .NOTES
                One table is output per table row with the -List parameter.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()] [System.Object] $Table
            )
            process {

                [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder';
                if ($Table.List) {

                    ## Create a table for each row
                    for ($r = 0; $r -lt $Table.Rows.Count; $r++) {

                        $row = $Table.Rows[$r];
                        if ($r -gt 0) {

                            ## Add a space between each table to mirror Word output rendering
                            [ref] $null = $tableBuilder.AppendLine('<p />');
                        }
                        [ref] $null = $tableBuilder.Append((GetHtmlTableList -Table $Table -Row $row));

                    } #end foreach row
                }
                else {

                    [ref] $null = $tableBuilder.Append((GetHtmlTable -Table $Table));
                } #end if
                return $tableBuilder.ToString();
                #Write-Output ($tableBuilder.ToString()) -NoEnumerate;

            } #end process
        } #end function outhtmltable


        function OutHtmlLineBreak {
        <#
            .SYNOPSIS
                Output formatted Html line break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {

                return '<hr />';

            }
        } #end function OutHtmlLineBreak


        function OutHtmlPageBreak {
        <#
            .SYNOPSIS
                Output formatted Html page break.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {

                [System.Text.StringBuilder] $pageBreakBuilder = New-Object 'System.Text.StringBuilder';
                [ref] $null = $pageBreakBuilder.Append('</div></page>');
                $topMargin = ConvertMmToEm $Document.Options['MarginTop'];
                $leftMargin = ConvertMmToEm $Document.Options['MarginLeft'];
                $bottomMargin = ConvertMmToEm $Document.Options['MarginBottom'];
                $rightMargin = ConvertMmToEm $Document.Options['MarginRight'];
                [ref] $null = $pageBreakBuilder.AppendFormat('<page><div class="{0}" style="padding-top: {1}em; padding-left: {2}em; padding-bottom: {3}em; padding-right: {4}em;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
                return $pageBreakBuilder.ToString();

            }
        } #end function OutHtmlPageBreak

        #endregion OutHtml Private Functions

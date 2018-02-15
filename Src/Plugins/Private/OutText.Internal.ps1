        #region OutText Private Functions

        function New-PScriboTextOption {
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
                ## Text/output width. 0 = none/no wrap.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $TextWidth = 120,

                ## Document header separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $HeaderSeparator = '=',

                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $SectionSeparator = '-',

                ## Document section separator character.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateLength(1,1)]
                [System.String] $LineBreakSeparator = '_',

                ## Default header/section separator width.
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Int32] $SeparatorWidth = $TextWidth,

                ## Text encoding
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
                [System.String] $Encoding = 'ASCII'
            )
            process {

                return @{
                    TextWidth = $TextWidth;
                    HeaderSeparator = $HeaderSeparator;
                    SectionSeparator = $SectionSeparator;
                    LineBreakSeparator = $LineBreakSeparator;
                    SeparatorWidth = $SeparatorWidth;
                    Encoding = $Encoding;
                }

            } #end process
        } #end function New-PScriboTextOption


        function OutTextTOC {
        <#
            .SYNOPSIS
                Output formatted Table of Contents
        #>
            [CmdletBinding()]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $TOC
            )
            begin {

                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {

                    $options = New-PScriboTextOption;
                }

            }
            process {

                $tocBuilder = New-Object -TypeName System.Text.StringBuilder;
                [ref] $null = $tocBuilder.AppendLine($TOC.Name);
                [ref] $null = $tocBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));

                if ($Options.ContainsKey('EnableSectionNumbering')) {

                    $maxSectionNumberLength = ([System.String] ($Document.TOC.Number | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum)).Length;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberPaddingLength = $maxSectionNumberLength - $tocEntry.Number.Length;
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        $sectionPadding = ''.PadRight($sectionNumberPaddingLength, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}  {2}{3}', $tocEntry.Number, $sectionPadding, $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }
                else {

                    $maxSectionNumberLength = $Document.TOC.Level | Sort-Object | Select-Object -Last 1;
                    foreach ($tocEntry in $Document.TOC) {
                        $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ');
                        [ref] $null = $tocBuilder.AppendFormat('{0}{1}', $sectionNumberIndent, $tocEntry.Name).AppendLine();
                    } #end foreach TOC entry
                }

                return $tocBuilder.ToString();

            } #end process
        } #end function OutTextTOC


        function OutTextBlankLine {
        <#
            .SYNOPSIS
                Output formatted text blankline.
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
                    [ref] $null = $blankLineBuilder.AppendLine();
                }
                return $blankLineBuilder.ToString();

            } #end process
        } #end function OutHtmlBlankLine


        function OutTextSection {
        <#
            .SYNOPSIS
                Output formatted text section.
        #>
            [CmdletBinding()]
            param (
                ## Section to output
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            begin {

                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {

                    $options = New-PScriboTextOption;
                }

            }
            process {

                $sectionBuilder = New-Object -TypeName System.Text.StringBuilder;
                if ($Document.Options['EnableSectionNumbering']) { [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name; }
                else { [string] $sectionName = '{0}' -f $Section.Name; }
                [ref] $null = $sectionBuilder.AppendLine();
                [ref] $null = $sectionBuilder.AppendLine($sectionName.TrimStart());
                [ref] $null = $sectionBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator));
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append(($s | OutTextParagraph)); }
                        'PScribo.PageBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextPageBreak)); }  ## Page breaks implemented as line break with extra padding
                        'PScribo.LineBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextLineBreak)); }
                        'PScribo.Table' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextTable)); }
                        'PScribo.BlankLine' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextBlankLine)); }
                        Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
                    } #end switch
                } #end foreach
                return $sectionBuilder.ToString();

            } #end process
        } #end function outtextsection


        function OutTextParagraph {
        <#
            .SYNOPSIS
                Output formatted paragraph text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            begin {

                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {

                    $options = New-PScriboTextOption;
                }

            }
            process {

                $padding = ''.PadRight(($Paragraph.Tabs * 4), ' ');
                if ([string]::IsNullOrEmpty($Paragraph.Text)) { $text = "$padding$($Paragraph.Id)"; }
                else { $text = "$padding$($Paragraph.Text)"; }

                $formattedText = OutStringWrap -InputObject $text -Width $Options.TextWidth;

                if ($Paragraph.NewLine) { return "$formattedText`r`n"; }
                else { return $formattedText; }

            } #end process
        } #end outtextparagraph


        function OutTextLineBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            begin {

                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {

                    $options = New-PScriboTextOption;
                }

            }
            process {

                ## Use the specified output width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                $lb = ''.PadRight($options.SeparatorWidth, $options.LineBreakSeparator);
                return "$(OutStringWrap -InputObject $lb -Width $options.TextWidth)`r`n";

            } #end process
        } #end function OutTextLineBreak


        function OutTextPageBreak {
        <#
            .SYNOPSIS
                Output formatted line break text.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param ( )
            process {
                return "$(OutTextLineBreak)`r`n";
            } #end process
        } #end function OutTextLineBreak


        function OutTextTable {
        <#
            .SYNOPSIS
                Output formatted text table.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Table
            )
            begin {

                ## Fix Set-StrictMode
                if (-not (Test-Path -Path Variable:\Options)) {

                    $options = New-PScriboTextOption;
                }

            }
            process {

                ## Use the current output buffer width
                if ($options.TextWidth -eq 0) { $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1; }
                if ($Table.List) {
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-List | Out-String -Width $options.TextWidth).Trim();
                } else {
                    ## Don't trim tabs for table headers
                    ## Tables set to AutoSize as otherwise, rendering is different between PoSh v4 and v5
                    $text = ($Table.Rows | Select-Object -Property * -ExcludeProperty '*__Style' | Format-Table -Wrap -AutoSize | Out-String -Width $options.TextWidth).Trim("`r`n");
                }
                # Ensure there's a space before and after the table.
                return "`r`n$text`r`n";

            } #end process
        } #end function outtexttable


        function OutStringWrap {
        <#
            .SYNOPSIS
                Outputs objects to strings, wrapping as required.
        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [Object[]] $InputObject,

                [Parameter()]
                [ValidateNotNull()]
                [System.Int32] $Width = $Host.UI.RawUI.BufferSize.Width
            )
            begin {

                ## 2 is the minimum, therefore default to wiiiiiiiiiide!
                if ($Width -lt 2) { $Width = 4096; }
                WriteLog -Message ('Wrapping text at "{0}" characters.' -f $Width) -IsDebug;

            }
            process {

                foreach ($object in $InputObject) {
                    $textBuilder = New-Object -TypeName System.Text.StringBuilder;
                    $text = (Out-String -InputObject $object).TrimEnd("`r`n");
                    for ($i = 0; $i -le $text.Length; $i += $Width) {
                        if (($i + $Width) -ge ($text.Length -1)) { [ref] $null = $textBuilder.Append($text.Substring($i)); }
                        else { [ref] $null = $textBuilder.AppendLine($text.Substring($i, $Width)); }
                    } #end for
                    return $textBuilder.ToString();
                    $textBuilder = $null;
                } #end foreach

            } #end process
        } #end function OutStringWrap

        #endregion OutText Private Functions

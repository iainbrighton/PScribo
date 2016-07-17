        #region TableStyle Private Functions
        function Add-PScriboTableStyle {
        <#
            .SYNOPSIS
                Defines a new PScribo table formatting style.
            .DESCRIPTION
                Creates a standard table formatting style that can be applied
                to the PScribo table keyword, e.g. a combination of header and
                row styles and borders.
            .NOTES
                Not all plugins support all options.
        #>
            [CmdletBinding()]
            param (
                ## Table Style name/id
                [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [Alias('Name')]
                [System.String] $Id,

                ## Header Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
                [ValidateNotNullOrEmpty()]
                [System.String] $HeaderStyle = 'Normal',

                ## Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
                [ValidateNotNullOrEmpty()]
                [System.String] $RowStyle = 'Normal',

                ## Header Row Style Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
                [AllowNull()]
                [Alias('AlternatingRowStyle')]
                [System.String] $AlternateRowStyle = 'Normal',

                ## Table border size/width (pt)
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
                [AllowNull()]
                [System.Single] $BorderWidth = 0,

                ## Table border colour
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Border')]
                [ValidateNotNullOrEmpty()]
                [Alias('BorderColour')]
                [System.String] $BorderColor = '000',

                ## Table cell top padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingTop = 1.0,

                ## Table cell left padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingLeft = 4.0,

                ## Table cell bottom padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingBottom = 0.0,

                ## Table cell right padding (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNull()]
                [System.Single] $PaddingRight = 4.0,

                ## Table alignment
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('Left','Center','Right')]
                [System.String] $Align = 'Left',

                ## Set as default table style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Default
            ) #end param
            begin {

                if ($BorderWidth -gt 0) { $borderStyle = 'Solid'; } else {$borderStyle = 'None'; }
                if (-not ($pscriboDocument.Styles.ContainsKey($HeaderStyle))) {
                    throw ($localized.UndefinedTableHeaderStyleError -f $HeaderStyle);
                }
                if (-not ($pscriboDocument.Styles.ContainsKey($RowStyle))) {
                    throw ($localized.UndefinedTableRowStyleError -f $RowStyle);
                }
                if (-not ($pscriboDocument.Styles.ContainsKey($AlternateRowStyle))) {
                    throw ($localized.UndefinedAltTableRowStyleError -f $AlternateRowStyle);
                }
                if (-not (Test-PScriboStyleColor -Color $BorderColor)) {
                    throw ($localized.InvalidTableBorderColorError -f $BorderColor);
                }

            } #end begin
            process {

                $pscriboDocument.Properties['TableStyles']++;
                $tableStyle = [PSCustomObject] @{
                    Id = $Id.Replace(' ', $pscriboDocument.Options['SpaceSeparator']);
                    Name = $Id;
                    HeaderStyle = $HeaderStyle;
                    RowStyle = $RowStyle;
                    AlternateRowStyle = $AlternateRowStyle;
                    PaddingTop = ConvertPtToMm $PaddingTop;
                    PaddingLeft = ConvertPtToMm $PaddingLeft;
                    PaddingBottom = ConvertPtToMm $PaddingBottom;
                    PaddingRight = ConvertPtToMm $PaddingRight;
                    Align = $Align;
                    BorderWidth = ConvertPtToMm $BorderWidth;
                    BorderStyle = $borderStyle;
                    BorderColor = Resolve-PScriboStyleColor -Color $BorderColor;
                }
                $pscriboDocument.TableStyles[$Id] = $tableStyle;
                if ($Default) { $pscriboDocument.DefaultTableStyle = $tableStyle.Id; }

            } #end process
        } #end function Add-PScriboTableStyle

        #endregion TableStyle Private Functions

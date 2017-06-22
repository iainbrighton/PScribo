        #region Style Private Functions

        function Add-PScriboStyle {
        <#
            .SYNOPSIS
                Initializes a new PScribo style object.
        #>
            [CmdletBinding()]
            param (
                ## Style name
                [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,

                ## Style id
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = $Name -Replace(' ',''),

                ## Font size (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.UInt16] $Size = 11,

                ## Font name (array of names for HTML output)
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String[]] $Font,

                ## Font color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Colour')]
                [ValidateNotNullOrEmpty()]
                [System.String] $Color = 'Black',

                ## Background color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('BackgroundColour')]
                [ValidateNotNullOrEmpty()]
                [System.String] $BackgroundColor,

                ## Bold typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Bold,

                ## Italic typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Italic,

                ## Underline typeface
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Underline,

                ## Text alignment
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateSet('Left','Center','Right','Justify')]
                [string] $Align = 'Left',

                ## Html CSS class id - to override Style.Id in HTML output.
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String] $ClassId = $Id,

                ## Hide style from UI (Word)
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Hide')]
                [System.Management.Automation.SwitchParameter] $Hidden,

                ## Set as default style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Default
            ) #end param
            begin {

                if (-not (Test-PScriboStyleColor -Color $Color)) {

                    throw ($localized.InvalidHtmlColorError -f $Color);
                }
                if ($BackgroundColor) {

                    if (-not (Test-PScriboStyleColor -Color $BackgroundColor)) {

                        throw ($localized.InvalidHtmlBackgroundColorError -f $BackgroundColor);
                    }
                    else {

                        $BackgroundColor = Resolve-PScriboStyleColor -Color $BackgroundColor;
                    }
                }
                if (-not ($Font)) {

                    $Font = $pscriboDocument.Options['DefaultFont'];
                }

            } #end begin
            process {

                $pscriboDocument.Properties['Styles']++;
                $style = [PSCustomObject] @{
                    Id   = $Id;
                    Name = $Name;
                    Font = $Font;
                    Size = $Size;
                    Color = (Resolve-PScriboStyleColor -Color $Color).ToLower();
                    BackgroundColor = $BackgroundColor.ToLower();
                    Bold = $Bold.ToBool();
                    Italic = $Italic.ToBool();
                    Underline = $Underline.ToBool();
                    Align = $Align;
                    ClassId = $ClassId;
                    Hidden = $Hidden.ToBool();
                }
                $pscriboDocument.Styles[$Id] = $style;
                if ($Default) { $pscriboDocument.DefaultStyle = $style.Id; }

            } #end process
        } #end function Add-PScriboStyle

        #endregion Style Private Functions

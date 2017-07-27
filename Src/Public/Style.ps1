function Resolve-PScriboStyleColor {
<#
    .SYNOPSIS
        Resolves a HTML color format or Word color constant to a RGB value
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNull()]
        [System.String] $Color
    )
    begin {

        # http://www.jadecat.com/tuts/colorsplus.html
        $wordColorConstants = @{
            AliceBlue = 'F0F8FF'; AntiqueWhite = 'FAEBD7'; Aqua = '00FFFF'; Aquamarine = '7FFFD4'; Azure = 'F0FFFF'; Beige = 'F5F5DC';
            Bisque = 'FFE4C4'; Black = '000000'; BlanchedAlmond = 'FFEBCD'; Blue = '0000FF'; BlueViolet = '8A2BE2'; Brown = 'A52A2A';
            BurlyWood = 'DEB887'; CadetBlue = '5F9EA0'; Chartreuse = '7FFF00'; Chocolate = 'D2691E'; Coral = 'FF7F50';
            CornflowerBlue = '6495ED'; Cornsilk = 'FFF8DC'; Crimson = 'DC143C'; Cyan = '00FFFF'; DarkBlue = '00008B'; DarkCyan = '008B8B';
            DarkGoldenrod = 'B8860B'; DarkGray = 'A9A9A9'; DarkGreen = '006400'; DarkKhaki = 'BDB76B'; DarkMagenta = '8B008B';
            DarkOliveGreen = '556B2F'; DarkOrange = 'FF8C00'; DarkOrchid = '9932CC'; DarkRed = '8B0000'; DarkSalmon = 'E9967A';
            DarkSeaGreen = '8FBC8F'; DarkSlateBlue = '483D8B'; DarkSlateGray = '2F4F4F'; DarkTurquoise = '00CED1'; DarkViolet = '9400D3';
            DeepPink = 'FF1493'; DeepSkyBlue = '00BFFF'; DimGray = '696969'; DodgerBlue = '1E90FF'; Firebrick = 'B22222';
            FloralWhite = 'FFFAF0'; ForestGreen = '228B22'; Fuchsia = 'FF00FF'; Gainsboro = 'DCDCDC'; GhostWhite = 'F8F8FF';
            Gold = 'FFD700'; Goldenrod = 'DAA520'; Gray = '808080'; Green = '008000'; GreenYellow = 'ADFF2F'; Honeydew = 'F0FFF0';
            HotPink = 'FF69B4'; IndianRed = 'CD5C5C'; Indigo = '4B0082'; Ivory = 'FFFFF0'; Khaki = 'F0E68C'; Lavender = 'E6E6FA';
            LavenderBlush = 'FFF0F5'; LawnGreen = '7CFC00'; LemonChiffon = 'FFFACD'; LightBlue = 'ADD8E6'; LightCoral = 'F08080';
            LightCyan = 'E0FFFF'; LightGoldenrodYellow = 'FAFAD2'; LightGreen = '90EE90'; LightGrey = 'D3D3D3'; LightPink = 'FFB6C1';
            LightSalmon = 'FFA07A'; LightSeaGreen = '20B2AA'; LightSkyBlue = '87CEFA'; LightSlateGray = '778899'; LightSteelBlue = 'B0C4DE';
            LightYellow = 'FFFFE0'; Lime = '00FF00'; LimeGreen = '32CD32'; Linen = 'FAF0E6'; Magenta = 'FF00FF'; Maroon = '800000';
            McMintGreen = 'BED6C9'; MediumAuqamarine = '66CDAA'; MediumBlue = '0000CD'; MediumOrchid = 'BA55D3'; MediumPurple = '9370D8';
            MediumSeaGreen = '3CB371'; MediumSlateBlue = '7B68EE'; MediumSpringGreen = '00FA9A'; MediumTurquoise = '48D1CC';
            MediumVioletRed = 'C71585'; MidnightBlue = '191970'; MintCream = 'F5FFFA'; MistyRose = 'FFE4E1'; Moccasin = 'FFE4B5';
            NavajoWhite = 'FFDEAD'; Navy = '000080'; OldLace = 'FDF5E6'; Olive = '808000'; OliveDrab = '688E23'; Orange = 'FFA500';
            OrangeRed = 'FF4500'; Orchid = 'DA70D6'; PaleGoldenRod = 'EEE8AA'; PaleGreen = '98FB98'; PaleTurquoise = 'AFEEEE';
            PaleVioletRed = 'D87093'; PapayaWhip = 'FFEFD5'; PeachPuff = 'FFDAB9'; Peru = 'CD853F'; Pink = 'FFC0CB'; Plum = 'DDA0DD';
            PowderBlue = 'B0E0E6'; Purple = '800080'; Red = 'FF0000'; RosyBrown = 'BC8F8F'; RoyalBlue = '4169E1'; SaddleBrown = '8B4513';
            Salmon = 'FA8072'; SandyBrown = 'F4A460'; SeaGreen = '2E8B57'; Seashell = 'FFF5EE'; Sienna = 'A0522D'; Silver = 'C0C0C0';
            SkyBlue = '87CEEB'; SlateBlue = '6A5ACD'; SlateGray = '708090'; Snow = 'FFFAFA'; SpringGreen = '00FF7F'; SteelBlue = '4682B4';
            Tan = 'D2B48C'; Teal = '008080'; Thistle = 'D8BFD8'; Tomato = 'FF6347'; Turquoise = '40E0D0'; Violet = 'EE82EE'; Wheat = 'F5DEB3';
            White = 'FFFFFF'; WhiteSmoke = 'F5F5F5'; Yellow = 'FFFF00'; YellowGreen = '9ACD32';
        };

    } #end begin
    process {

        $pscriboColor = $Color;
        if ($wordColorConstants.ContainsKey($pscriboColor)) {

            return $wordColorConstants[$pscriboColor].ToLower();
        }
        elseif ($pscriboColor.Length -eq 6 -or $pscriboColor.Length -eq 3) {

            $pscriboColor = '#{0}' -f $pscriboColor;
        }
        elseif ($pscriboColor.Length -eq 7 -or $pscriboColor.Length -eq 4) {

            if (-not ($pscriboColor.StartsWith('#'))) {

                return $null;
            }
        }
        if ($pscriboColor -notmatch '^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$') {

            return $null;
        }
        return $pscriboColor.TrimStart('#').ToLower();

    } #end process
} #end function ResolvePScriboColor


function Test-PScriboStyleColor {
<#
    .SYNOPSIS
        Tests whether a color string is a valid HTML color.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color
    )
    process {

        if (Resolve-PScriboStyleColor -Color $Color) { return $true; }
        else { return $false; }

    } #end process
} #end function test-pscribostylecolor


function Test-PScriboStyle {
<#
    .SYNOPSIS
        Tests whether a style has been defined.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name
    )
    process {

        return $PScriboDocument.Styles.ContainsKey($Name);

    }
} #end function Test-PScriboStyle


function Style {
<#
    .SYNOPSIS
        Defines a new PScribo formatting style.
    .DESCRIPTION
        Creates a standard format formatting style that can be applied
        to PScribo document keywords, e.g. a combination of font style, font
        weight and font size.
    .NOTES
        Not all plugins support all options.
#>
    [CmdletBinding()]
    param (
        ## Style name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## Font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [System.UInt16] $Size = 11,

        ## Font color/colour
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Colour')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Color = '000',

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
        [System.String] $Align = 'Left',

        ## Set as default style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Default,

        ## Style id
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = $Name -Replace(' ',''),

        ## Font name (array of names for HTML output)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $Font,

        ## Html CSS class id - to override Style.Id in HTML output.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $ClassId = $Id,

        ## Hide style from UI (Word)
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Hide')]
        [System.Management.Automation.SwitchParameter] $Hidden
    )
    begin {

        <#! Style.Internal.ps1 !#>

    }
    process {

        WriteLog -Message ($localized.ProcessingStyle -f $Id);
        Add-PScriboStyle @PSBoundParameters;

    } #end process
} #end function Style


function Set-Style {
<#
    .SYNOPSIS
        Sets the style for an individual table row or cell.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Object])]
    param (
        ## PSCustomObject to apply the style to
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object[]] [Ref] $InputObject,

        ## PScribo style Id to apply
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Property name(s) to apply the selected style to. Leave blank to apply the style to the entire row.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Property = '',

        ## Passes the modified object back to the pipeline
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    ) #end param
    begin {

        if (-not (Test-PScriboStyle -Name $Style)) {
            Write-Error ($localized.UndefinedStyleError -f $Style);
            return;
        }

    }
    process {

        foreach ($object in $InputObject) {

            foreach ($p in $Property) {

                ## If $Property not set, __Style will apply to the whole row.
                $propertyName = '{0}__Style' -f $p;
                $object | Add-Member -MemberType NoteProperty -Name $propertyName -Value $Style -Force;
            }
        }
        if ($PassThru) {

            return $object;
        }

    } #end process
} #end function Set-Style

function Add-PScriboTableStyle
{
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
    param
    (
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

        ## Alternating Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [AllowNull()]
        [Alias('AlternatingRowStyle')]
        [System.String] $AlternateRowStyle = $RowStyle,

        ## Caption Style Id
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $CaptionStyle = 'Normal',

        ## Table border size/width (pt)
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Single] $BorderWidth = 0,

        ## Table border colour
        [Parameter(ValueFromPipelineByPropertyName)]
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

        ## Caption prefix
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $CaptionPrefix = 'Table',

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Above', 'Below')]
        [System.String] $CaptionLocation = 'Below',

        ## Set as default table style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Default
    )
    begin
    {
        if ($BorderWidth -gt 0)
        {
            $borderStyle = 'Solid'
        }
        else
        {
            $borderStyle = 'None'
        }
        if (-not ($pscriboDocument.Styles.ContainsKey($HeaderStyle)))
        {
            throw ($localized.UndefinedTableHeaderStyleError -f $HeaderStyle)
        }
        if (-not ($pscriboDocument.Styles.ContainsKey($RowStyle)))
        {
            throw ($localized.UndefinedTableRowStyleError -f $RowStyle)
        }
        if (-not ($pscriboDocument.Styles.ContainsKey($AlternateRowStyle)))
        {
            throw ($localized.UndefinedAltTableRowStyleError -f $AlternateRowStyle)
        }
        if (-not (Test-PScriboStyleColor -Color $BorderColor))
        {
            throw ($localized.InvalidTableBorderColorError -f $BorderColor)
        }
    }
    process
    {
        $pscriboDocument.Properties['TableStyles']++;
        $tableStyle = [PSCustomObject] @{
            Id                = $Id.Replace(' ', $pscriboDocument.Options['SpaceSeparator'])
            Name              = $Id
            HeaderStyle       = $HeaderStyle
            RowStyle          = $RowStyle
            AlternateRowStyle = $AlternateRowStyle
            CaptionStyle      = $CaptionStyle
            PaddingTop        = ConvertTo-Mm -Point $PaddingTop
            PaddingLeft       = ConvertTo-Mm -Point $PaddingLeft
            PaddingBottom     = ConvertTo-Mm -Point $PaddingBottom
            PaddingRight      = ConvertTo-Mm -Point $PaddingRight
            Align             = $Align
            BorderWidth       = ConvertTo-Mm -Point $BorderWidth
            BorderStyle       = $borderStyle
            BorderColor       = Resolve-PScriboStyleColor -Color $BorderColor
            CaptionPrefix     = $CaptionPrefix
            CaptionLocation   = $CaptionLocation
        }
        $pscriboDocument.TableStyles[$Id] = $tableStyle
        if ($Default)
        {
            $pscriboDocument.DefaultTableStyle = $tableStyle.Id
        }
    }
}

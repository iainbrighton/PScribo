function TableStyle
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
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param
    (
        ## Table Style name/id
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Padding')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String] $Id,

        ## Header Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [System.String] $HeaderStyle = 'Default',

        ## Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 2, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, Position = 2, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [System.String] $RowStyle = 'Default',

        ## Alternating Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 3, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, Position = 3, ParameterSetName = 'Default')]
        [AllowNull()]
        [Alias('AlternatingRowStyle')]
        [System.String] $AlternateRowStyle = $RowStyle,

        ## Table border size/width (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [AllowNull()]
        [System.Single] $BorderWidth = 0,

        ## Table border colour
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [Alias('BorderColour')]
        [System.String] $BorderColor = '000',

        ## Table cell padding (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [ValidateNotNull()]
        [System.Single] $Padding,

        ## Table cell top padding (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [ValidateNotNull()]
        [System.Single] $PaddingTop = 1.0,

        ## Table cell left padding (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [ValidateNotNull()]
        [System.Single] $PaddingLeft = 4.0,

        ## Table cell bottom padding (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [ValidateNotNull()]
        [System.Single] $PaddingBottom = 0.0,

        ## Table cell right padding (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [ValidateNotNull()]
        [System.Single] $PaddingRight = 4.0,

        ## Table alignment
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [ValidateSet('Left','Center','Right')]
        [System.String] $Align = 'Left',

        ## Table caption prefix
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.String] $CaptionPrefix = 'Table',

        ## Table caption prefix
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.String] $CaptionStyle = 'Caption',

        ## Table caption display location.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [ValidateSet('Above', 'Below')]
        [System.String] $CaptionLocation = 'Below',

        ## Set as default table style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Padding')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Default')]
        [System.Management.Automation.SwitchParameter] $Default
    )
    process
    {
        if ($PSBoundParameters.ContainsKey('Padding'))
        {
            $PSBoundParameters['PaddingTop'] = $Padding
            $PSBoundParameters['PaddingLeft'] = $Padding
            $PSBoundParameters['PaddingBottom'] = $Padding
            $PSBoundParameters['PaddingRight'] = $Padding
            $null = $PSBoundParameters.Remove('Padding')
        }
        Write-PScriboMessage -Message ($localized.ProcessingTableStyle -f $Id)
        Add-PScriboTableStyle @PSBoundParameters
    }
}

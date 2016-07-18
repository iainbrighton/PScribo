function TableStyle {
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
        [System.String] $HeaderStyle = 'Default',

        ## Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String] $RowStyle = 'Default',

        ## Header Row Style Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [AllowNull()]
        [Alias('AlternatingRowStyle')]
        [System.String] $AlternateRowStyle = 'Default',

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
        <#! TableStyle.Internal.ps1 !#>
    }
    process {

        WriteLog -Message ($localized.ProcessingTableStyle -f $Id);
        Add-PScriboTableStyle @PSBoundParameters;

    }
} #end function tablestyle

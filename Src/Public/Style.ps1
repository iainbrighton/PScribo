
function Style
{
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
    param
    (
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
    process
    {
        Write-PScriboMessage -Message ($localized.ProcessingStyle -f $Id)
        Add-PScriboStyle @PSBoundParameters
    }
}

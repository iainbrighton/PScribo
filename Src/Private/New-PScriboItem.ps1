function New-PScriboItem
{
<#
    .SYNOPSIS
        Initializes new PScribo list item object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.String] $Text,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Style')]
        [System.String] $Style,

        ## Override the bold style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Bold,

        ## Override the italic style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Italic,

        ## Override the underline style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Underline,

        ## Override the font name(s)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Font,

        ## Override the font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [AllowNull()]
        [System.UInt16] $Size = $null,

        ## Override the font color/colour
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [AllowNull()]
        [System.String] $Color = $null
    )
    process
    {
        $pscriboItem = [PSCustomObject] @{
            Id               = [System.Guid]::NewGuid().ToString()
            Level            = 0
            Index            = 0
            Number           = ''
            Text             = $Text
            Type             = 'PScribo.Item'
            Style            = $Style
            Bold             = $Bold
            Italic           = $Italic
            Underline        = $Underline
            Font             = $Font
            Size             = $Size
            Color            = $Color
            IsStyleInherited = $PSCmdlet.ParameterSetName -eq 'Default'
            HasStyle         = $PSCmdlet.ParameterSetName -eq 'Style'
            HasInlineStyle   = $PSCmdlet.ParameterSetName -eq 'Inline'
        }
        return $pscriboItem
    }
}

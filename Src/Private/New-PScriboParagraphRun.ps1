function New-PScriboParagraphRun
{
<#
    .SYNOPSIS
        Initializes a new PScribo paragraph run (text) object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [AllowEmptyString()]
        [System.String] $Text,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Style')]
        [System.String] $Style,

        ## No space applied between this text block and the next
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoSpace,

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
        $pscriboParagraphRun = [PSCustomObject] @{
            Type              = 'PScribo.ParagraphRun'
            Text              = $Text
            Style             = $Style
            NoSpace           = $NoSpace
            Bold              = $Bold
            Italic            = $Italic
            Underline         = $Underline
            Font              = $Font
            Size              = $Size
            Color             = $Color
            IsStyleInherited  = $PSCmdlet.ParameterSetName -eq 'Default'
            HasStyle          = $PSCmdlet.ParameterSetName -eq 'Style'
            HasInlineStyle    = $PSCmdlet.ParameterSetName -eq 'Inline'
            IsParagraphRunEnd = $false
            Name              = $null # For legacy Xml output
            Value             = $null # For legacy Xml output
        }
        return $pscriboParagraphRun
    }
}

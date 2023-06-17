function NumberStyle
{
<#
    .SYNOPSIS
        Defines a new PScribo numbered list formatting style.

    .DESCRIPTION
        Creates a number list formatting style that can be applied to the PScribo 'List' keyword.

    .NOTES
        Not all plugins support all options.
#>
    [CmdletBinding(DefaultParameterSetName = 'Predefined')]
    param
    (
        ## Table Style name/id
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Predefined')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Custom')]
        [ValidateNotNullOrEmpty()]
        [Alias('Name')]
        [System.String] $Id,

        ## NOTE: Only supported in Text, Html and Word output.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [ValidateSet('Number','Letter','Roman')]
        [System.String] $Format,

        ## Custom number 'XYZ-###' NOTE: Only supported in Text and Word output.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Custom')]
        [System.String] $Custom,

        ## Number format suffix, e.g. '.' or ')'. NOTE: Only supported in text and Word output
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [ValidateLength(1, 1)]
        [System.String] $Suffix = '.',

        ## Only applicable to 'Letter' and 'Roman' formats
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [System.Management.Automation.SwitchParameter] $Uppercase,

        ## Set as default table style. NOTE: Cannot set custom styles as default as they're not supported by all plugins.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [System.Management.Automation.SwitchParameter] $Default,

        ## Number alignment.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Custom')]
        [ValidateSet('Left', 'Right')]
        [System.String] $Align = 'Right',

        ## Override the default Word indentation level.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Custom')]
        [System.Int32] $Indent,

        ## Override the default Word hanging indentation level.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Predefined')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Custom')]
        [System.Int32] $Hanging
    )
    process
    {
        Write-PScriboMessage -Message ($localized.ProcessingNumberStyle -f $Id)
        Add-PScriboNumberStyle @PSBoundParameters
    }
}

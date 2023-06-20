function Item
{
<#
    .SYNOPSIS
        Initializes a new PScribo list Item object.
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## List item text.
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.String] $Text,

        ## List item style Name/Id reference.
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
    begin
    {
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($psCallStack[1].FunctionName -ne 'List<Process>')
        {
            throw $localized.ItemRootError
        }
    }
    process
    {
        return (New-PScriboItem @PSBoundParameters)
    }
}

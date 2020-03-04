function New-PScriboTable
{
<#
    .SYNOPSIS
        Initializes a new PScribo table object.

    .PARAMETER Name

    .PARAMETER Columns

    .PARAMETER ColumnWidths

    .PARAMETER Rows

    .PARAMETER Style

    .PARAMETER Width

    .PARAMETER List

    .PARAMETER Width

    .PARAMETER Tabs

    .PARAMETER Caption

    .PARAMETER CaptionTop

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## Table name/Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name = ([System.Guid]::NewGuid().ToString()),

        ## Table columns/display order
        [Parameter(Mandatory)]
        [AllowNull()]
        [System.String[]] $Columns,

        ## Table columns widths
        [Parameter(Mandatory)]
        [AllowNull()]
        [System.UInt16[]] $ColumnWidths,

        ## Collection of PScriboTableObjects for table rows
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Collections.ArrayList] $Rows,

        ## Table style
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Style,

        # List view
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
        [System.Management.Automation.SwitchParameter] $List,

        ## Combine list view based upon the specified key/property name
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
        [System.String] $ListKey = $null,

        ## Display the key name in the table output
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
        [System.Management.Automation.SwitchParameter] $DisplayListKey,

        ## Table width (%), 0 = Autofit
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 100)]
        [System.UInt16] $Width = 100,

        ## Indent table
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 10)]
        [System.UInt16] $Tabs,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Caption
    )
    process
    {
        $typeName = 'PScribo.Table'
        $pscriboDocument.Properties['Tables']++
        $pscriboTable = [PSCustomObject] @{
            Id              = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper()
            Name            = $Name
            Number          = $pscriboDocument.Properties['Tables']
            Type            = $typeName
            Columns         = $Columns
            ColumnWidths    = $ColumnWidths
            Rows            = $Rows
            Style           = $Style
            Width           = $Width
            Tabs            = $Tabs
            HasCaption      = $PSBoundParameters.ContainsKey('Caption')
            Caption         = $Caption
            IsList          = $List
            IsKeyedList     = $PSBoundParameters.ContainsKey('ListKey')
            ListKey         = $ListKey
            DisplayListKey  = $DisplayListKey.ToBool()
        }
        return $pscriboTable
    }
}

function New-PScriboFormattedTableRowCell
{
<#
    .SYNOPSIS
        Creates a formatted table cell for plugin output/rendering.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSObject])]
    param
    (
        [Parameter(ValueFromPipeline)]
        [AllowNull()]
        [System.String] $Content,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Uint16] $Width,

        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.String] $Style = $null,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Boolean] $IsAlternateRow
    )
    process
    {
        $isStyleInherited = $true

        if (-not ([System.String]::IsNullOrEmpty($Style)))
        {
            ## Use the explictit style
            $isStyleInherited = $false
        }

        return [PSCustomObject] @{
            Content          = $Content;
            Width            = $Width;
            Style            = $Style;
            IsStyleInherited = $isStyleInherited;
        }
    }
}

function New-PScriboHtmlOption
{
<#
    .SYNOPSIS
        Sets the text plugin specific formatting/output options.
    .NOTES
        All plugin options should be prefixed with the plugin name.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Boolean] $NoPageLayoutStyle = $false
    )
    process
    {
        return @{
            NoPageLayoutStyle = $NoPageLayoutStyle
        }
    }
}

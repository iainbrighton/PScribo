function New-PScriboJsonOption
{
<#
    .SYNOPSIS
        Sets the Json plugin specific formatting/output options.

    .NOTES
        All plugin options should be prefixed with the plugin name.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        ## Text encoding
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
        [System.String] $Encoding = 'ASCII'
    )
    process
    {
        return @{
            Encoding = $Encoding
        }
    }
}

function Merge-PScriboPluginOption
{
<#
    .SYNOPSIS
        Merges the specified options along with the plugin-specific default options.
#>
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        ## Default/document options
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $DocumentOptions,

        ## Default plugin options to merge
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Collections.Hashtable] $DefaultPluginOptions,

        ## Specified runtime plugin options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $PluginOptions
    )
    process
    {
        $mergedOptions = $DocumentOptions.Clone();

        if ($null -ne $DefaultPluginOptions)
        {
            ## Overwrite the default document option with the plugin default option/value
            foreach ($option in $DefaultPluginOptions.GetEnumerator())
            {
                $mergedOptions[$($option.Key)] = $option.Value;
            }
        }

        if ($null -ne $PluginOptions)
        {
            ## Overwrite the default document/plugin default option/value with the specified/runtime option
            foreach ($option in $PluginOptions.GetEnumerator())
            {
                $mergedOptions[$($option.Key)] = $option.Value;
            }
        }
        return $mergedOptions;
    }
}

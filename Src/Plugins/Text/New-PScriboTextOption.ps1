function New-PScriboTextOption
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
        ## Text/output width. 0 = none/no wrap.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Int32] $TextWidth = 120,

        ## Document header separator character.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateLength(1,1)]
        [System.String] $HeaderSeparator = '=',

        ## Document section separator character.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateLength(1,1)]
        [System.String] $SectionSeparator = '-',

        ## Document section separator character.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateLength(1,1)]
        [System.String] $LineBreakSeparator = '_',

        ## Default header/section separator width.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Int32] $SeparatorWidth = $TextWidth,

        ## Text encoding
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
        [System.String] $Encoding = 'ASCII'
    )
    process
    {
        return @{
            TextWidth = $TextWidth
            HeaderSeparator = $HeaderSeparator
            SectionSeparator = $SectionSeparator
            LineBreakSeparator = $LineBreakSeparator
            SeparatorWidth = $SeparatorWidth
            Encoding = $Encoding
        }
    }
}

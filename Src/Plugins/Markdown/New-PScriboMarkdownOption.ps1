function New-PScriboMarkdownOption
{
<#
    .SYNOPSIS
        Sets the Markdown plugin specific formatting/output options.

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

        ## Embed image data at bottom of the document.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Boolean] $EmbedImage = $false,

        ## Render blank lines as <br /> for Html page rendering.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Boolean] $RenderBlankline = $false,

        ## ocument page break separatorseparator character.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateLength(1,1)]
        [System.String] $PageBreakSeparator = '*',

        ## Document page break separator separator width.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Int32] $PageBreakSeparatorWidth = '20',

        ## Document line break separator character.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateLength(1,1)]
        [System.String] $LineBreakSeparator = '-',

        ## Document line break separator width.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.Int32] $LineBreakSeparatorWidth = '10',

        ## Text encoding
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('ASCII','Unicode','UTF7','UTF8')]
        [System.String] $Encoding = 'ASCII'
    )
    process
    {
        return @{
            TextWidth               = $TextWidth
            PageBreakSeparator      = $PageBreakSeparator
            PageBreakSeparatorWidth = $PageBreakSeparatorWidth
            LineBreakSeparator      = $LineBreakSeparator
            LineBreakSeparatorWidth = $LineBreakSeparatorWidth
            EmbedImage              = $EmbedImage
            RenderBlankline         = $RenderBlankline
            Encoding                = $Encoding
        }
    }
}

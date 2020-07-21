function Get-MarkdownInlineStyle
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [System.String] $Style
    )
    process
    {
        $isBold = $false
        $isItalic = $false

        $pscriboStyle = Get-PScriboDocumentStyle -Style $Style
        if ($null -ne $pscriboStyle.Bold)
        {
            $isBold = $pscriboStyle.Bold
        }
        if ($null -ne $pscriboStyle.Italic)
        {
            $isItalic = $pscriboStyle.Italic
        }

        if ($isBold -and $isItalic)
        {
            return '***'
        }
        elseif ($isBold)
        {
            return '**'
        }
        elseif ($isItalic)
        {
            return '_'
        }
    }
}

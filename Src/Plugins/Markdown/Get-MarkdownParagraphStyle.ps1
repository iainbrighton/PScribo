function Get-MarkdownParagraphStyle
{
<#
    .SYNOPSIS
        Generates Markdown paragraph styling markup.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Paragraph
    )
    process
    {
        $isBold = $false
        $isItalic = $false

        ## Evaluate the paragraph style
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $paragraphStyle = Get-PScriboDocumentStyle -Style $Paragraph.Style
            if ($null -ne $paragraphStyle.Bold)
            {
                $isBold = $paragraphStyle.Bold
            }
            if ($null -ne $paragraphStyle.Italic)
            {
                $isItalic = $paragraphStyle.Italic
            }
        }

        return (Get-MarkdownStyle -Bold:$isBold -Italic:$isItalic)
    }
}

function Get-MarkdownParagraphRunInlineStyle
{
<#
    .SYNOPSIS
        Generates Markdown paragraph run styling markup.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $ParagraphRun
    )
    process
    {
        $isBold = $false
        $isItalic = $false

        ## Evaluate the paragraph run style
        if ($ParagraphRun.HasStyle)
        {
            $paragraphRunStyle = Get-PScriboDocumentStyle -Style $ParagraphRun.Style
            if (($null -ne $paragraphRunStyle.Bold) -and ($paragraphRunStyle.Bold -ne $isBold))
            {
                $isBold = $paragraphRunStyle.Bold
            }
            if (($null -ne $paragraphRunStyle.Italic) -and ($paragraphRunStyle.Italic -ne $isItalic))
            {
                $isItalic = $paragraphRunStyle.Italic
            }
        }

        ## Finally evaluate any inline style
        if ($ParagraphRun.HasInlineStyle)
        {
            if (($null -ne $paragraphRun.Bold) -and ($paragraphRun.Bold -ne $isBold))
            {
                $isBold = $paragraphRun.Bold
            }
            if (($null -ne $paragraphRun.Italic) -and ($paragraphRun.Italic -ne $isItalic))
            {
                $isItalic = $paragraphRun.Italic
            }
        }

        return (Get-MarkdownStyle -Bold:$isBold -Italic:$isItalic)
    }
}

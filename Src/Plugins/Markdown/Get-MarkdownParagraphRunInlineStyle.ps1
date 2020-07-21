function Get-MarkdownParagraphRunInlineStyle
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $ParagraphRun,

        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [System.String] $Style
    )
    process
    {
        $isBold = $false
        $isItalic = $false

        ## Evaluate the paragraph style
        if (-not [System.String]::IsNullOrEmpty($Style))
        {
            $paragraphStyle = Get-PScriboDocumentStyle -Style $Style
            if ($null -ne $paragraphStyle.Bold)
            {
                $isBold = $paragraphStyle.Bold
            }
            if ($null -ne $paragraphStyle.Italic)
            {
                $isItalic = $paragraphStyle.Italic
            }
        }

        ## Override with the paragraph run style
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

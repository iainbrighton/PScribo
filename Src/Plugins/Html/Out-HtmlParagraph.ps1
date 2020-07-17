function Out-HtmlParagraph
{
<#
    .SYNOPSIS
        Output formatted Html paragraph run.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Paragraph
    )
    process
    {
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'

        if ([System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            if ($Paragraph.Tabs -gt 0)
            {
                ## Default to 1/2in tab spacing
                $tabEm = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Paragraph.Tabs)) -Format 'f2'
                [ref] $null = $paragraphBuilder.AppendFormat('<div style="margin-left: {0}rem;" >', $tabEm)
            }
            else
            {
                [ref] $null = $paragraphBuilder.Append('<div>')
            }
        }
        else
        {
            if ($Paragraph.Tabs -gt 0)
            {
                ## Default to 1/2in tab spacing
                $tabEm = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Paragraph.Tabs)) -Format 'f2'
                [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}" style="margin-left: {1}rem;" >', $Paragraph.Style, $tabEm)
            }
            else
            {
                [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">', $Paragraph.Style)
            }
        }

        foreach ($paragraphRun in $Paragraph.Sections)
        {
            if (($paragraphRun.HasStyle -eq $true) -and ($paragraphRun.HasInlineStyle -eq $true))
            {
                $inlineStyle = Get-HtmlParagraphInlineStyle -Paragraph $paragraphRun -NoIndent
                [ref] $null = $paragraphBuilder.AppendFormat('<span class="{0}" style="{1}">', $paragraphRun.Style, $inlineStyle)
            }
            elseif ($paragraphRun.HasStyle)
            {
                [ref] $null = $paragraphBuilder.AppendFormat('<span class="{0}">', $paragraphRun.Style)
            }
            elseif ($paragraphRun.HasInlineStyle)
            {
                $inlineStyle = Get-HtmlParagraphInlineStyle -Paragraph $paragraphRun -NoIndent
                [ref] $null = $paragraphBuilder.AppendFormat('<span style="{0}">', $inlineStyle)
            }

            $text = Resolve-PScriboToken -InputObject $paragraphRun.Text
            $encodedText = [System.Net.WebUtility]::HtmlEncode($text)
            $encodedText = $encodedText.Replace([System.Environment]::NewLine, '<br />')
            [ref] $null = $paragraphBuilder.Append($encodedText)

            if (($paragraphRun.IsParagraphRunEnd -eq $false) -and
                ($paragraphRun.NoSpace -eq $false))
            {
                [ref] $null = $paragraphBuilder.Append(' ')
            }

            if ($paragraphRun.HasStyle -or $paragraphRun.HasInlineStyle)
            {
                [ref] $null = $paragraphBuilder.Append('</span>')
            }
        }

        [ref] $null = $paragraphBuilder.Append('</div>')
        return $paragraphBuilder.ToString()
    }
}

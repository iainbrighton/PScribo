function Out-HtmlParagraph
{
<#
    .SYNOPSIS
        Output formatted Html paragraph.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Paragraph,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ParseToken
    )
    process
    {
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'

        $text = $Paragraph.Text
        if ([System.String]::IsNullOrEmpty($text))
        {
            $text = $Paragraph.Id
        }

        if ($ParseToken)
        {
            $text = Resolve-PScriboToken -InputObject $text
        }

        $encodedText = [System.Net.WebUtility]::HtmlEncode($text)
        $encodedText = $encodedText.Replace([System.Environment]::NewLine, '<br />')
        $customStyle = Get-HtmlParagraphStyle -Paragraph $Paragraph
        if ([System.String]::IsNullOrEmpty($Paragraph.Style) -and [System.String]::IsNullOrEmpty($customStyle))
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div>{0}</div>', $encodedText)
        }
        elseif ([System.String]::IsNullOrEmpty($customStyle))
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">{1}</div>', $Paragraph.Style, $encodedText)
        }
        else
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div style="{0}">{1}</div>', $customStyle, $encodedText)
        }
        return $paragraphBuilder.ToString()
    }
}

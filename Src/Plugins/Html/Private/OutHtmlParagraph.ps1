function OutHtmlParagraph
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
        [System.Management.Automation.PSObject] $Paragraph
    )
    process
    {
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder';
        $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Text);
        if ([System.String]::IsNullOrEmpty($encodedText))
        {
            $encodedText = [System.Net.WebUtility]::HtmlEncode($Paragraph.Id);
        }
        # $encodedText = $encodedText -replace [System.Environment]::NewLine, '<br />';
        $encodedText = $encodedText.Replace([System.Environment]::NewLine, '<br />');
        $customStyle = GetHtmlParagraphStyle -Paragraph $Paragraph;
        if ([System.String]::IsNullOrEmpty($Paragraph.Style) -and [System.String]::IsNullOrEmpty($customStyle))
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div>{0}</div>', $encodedText);
        }
        elseif ([System.String]::IsNullOrEmpty($customStyle))
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div class="{0}">{1}</div>', $Paragraph.Style, $encodedText);
        }
        else
        {
            [ref] $null = $paragraphBuilder.AppendFormat('<div style="{0}">{1}</div>', $customStyle, $encodedText);
        }
        return $paragraphBuilder.ToString();
    }
}

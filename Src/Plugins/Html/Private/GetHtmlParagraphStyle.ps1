function GetHtmlParagraphStyle
{
<#
    .SYNOPSIS
        Generates html style attribute from PScribo paragraph style overrides.
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
        $paragraphStyleBuilder = New-Object -TypeName System.Text.StringBuilder;
        if ($Paragraph.Tabs -gt 0)
        {
            ## Default to 1/2in tab spacing
            $tabEm = ConvertToInvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Paragraph.Tabs)) -Format 'f2';
            [ref] $null = $paragraphStyleBuilder.AppendFormat(' margin-left: {0}rem;', $tabEm);
        }
        if ($Paragraph.Font)
        {
            [ref] $null = $paragraphStyleBuilder.AppendFormat(" font-family: '{0}';", $Paragraph.Font -Join "','");
        }
        if ($Paragraph.Size -gt 0)
        {
            ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
            $invariantParagraphSize = ConvertToInvariantCultureString -Object ($Paragraph.Size / 12) -Format 'f2';
            [ref] $null = $paragraphStyleBuilder.AppendFormat(' font-size: {0}rem;', $invariantParagraphSize);
        }
        if ($Paragraph.Bold -eq $true)
        {
            [ref] $null = $paragraphStyleBuilder.Append(' font-weight: bold;');
        }
        if ($Paragraph.Italic -eq $true)
        {
            [ref] $null = $paragraphStyleBuilder.Append(' font-style: italic;');
        }
        if ($Paragraph.Underline -eq $true)
        {
            [ref] $null = $paragraphStyleBuilder.Append(' text-decoration: underline;');
        }
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Color) -and $Paragraph.Color.StartsWith('#'))
        {
            [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: {0};', $Paragraph.Color.ToLower());
        }
        elseif (-not [System.String]::IsNullOrEmpty($Paragraph.Color))
        {
            [ref] $null = $paragraphStyleBuilder.AppendFormat(' color: #{0};', $Paragraph.Color.ToLower());
        }
        return $paragraphStyleBuilder.ToString().TrimStart();
    }
}

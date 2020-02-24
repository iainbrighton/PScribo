function GetHtmlStyle
{
<#
    .SYNOPSIS
        Generates html stylesheet style attributes from a PScribo document style.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        ## PScribo document style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Style
    )
    process
    {
        $styleBuilder = New-Object -TypeName System.Text.StringBuilder;
        [ref] $null = $styleBuilder.AppendFormat(" font-family: '{0}';", $Style.Font -join "','");
        ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
        $invariantFontSize = ConvertTo-InvariantCultureString -Object ($Style.Size / 12) -Format 'f2';
        [ref] $null = $styleBuilder.AppendFormat(' font-size: {0}rem;', $invariantFontSize);
        [ref] $null = $styleBuilder.AppendFormat(' text-align: {0};', $Style.Align.ToLower());
        if ($Style.Bold) {

            [ref] $null = $styleBuilder.Append(' font-weight: bold;');
        }
        else {

            [ref] $null = $styleBuilder.Append(' font-weight: normal;');
        }
        if ($Style.Italic) {

            [ref] $null = $styleBuilder.Append(' font-style: italic;');
        }
        if ($Style.Underline) {

            [ref] $null = $styleBuilder.Append(' text-decoration: underline;');
        }
        if ($Style.Color.StartsWith('#')) {

            [ref] $null = $styleBuilder.AppendFormat(' color: {0};', $Style.Color.ToLower());
        }
        else {

            [ref] $null = $styleBuilder.AppendFormat(' color: #{0};', $Style.Color);
        }
        if ($Style.BackgroundColor) {

            if ($Style.BackgroundColor.StartsWith('#')) {

                [ref] $null = $styleBuilder.AppendFormat(' background-color: {0};', $Style.BackgroundColor.ToLower());
            }
            else {

                [ref] $null = $styleBuilder.AppendFormat(' background-color: #{0};', $Style.BackgroundColor.ToLower());
            }
        }
        return $styleBuilder.ToString();
    }
}

function Get-HtmlStyle
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
        $styleBuilder = New-Object -TypeName System.Text.StringBuilder
        [ref] $null = $styleBuilder.AppendFormat(" font-family: '{0}';", $Style.Font -join "','")
        ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
        $invariantFontSize = ConvertTo-InvariantCultureString -Object ($Style.Size / 12) -Format 'f2'
        [ref] $null = $styleBuilder.AppendFormat(' font-size: {0}rem;', $invariantFontSize)
        [ref] $null = $styleBuilder.AppendFormat(' text-align: {0};', $Style.Align.ToLower())

        if ($Style.Bold)
        {
            [ref] $null = $styleBuilder.Append(' font-weight: bold;')
        }
        else
        {
            [ref] $null = $styleBuilder.Append(' font-weight: normal;')
        }

        if ($Style.Italic)
        {
            [ref] $null = $styleBuilder.Append(' font-style: italic;')
        }

        if ($Style.Underline)
        {
            [ref] $null = $styleBuilder.Append(' text-decoration: underline;')
        }

        if ($Style.Color)
        {
            $color = Resolve-PScriboStyleColor -Color $Style.Color
            [ref] $null = $styleBuilder.AppendFormat(' color: #{0};', $color)
        }

        if ($Style.BackgroundColor)
        {
            $backgroundColor = Resolve-PScriboStyleColor -Color $Style.BackgroundColor
            [ref] $null = $styleBuilder.AppendFormat(' background-color: #{0};', $backgroundColor)
        }

        return $styleBuilder.ToString()
    }
}

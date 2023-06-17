function Get-HtmlListItemInlineStyle
{
<#
    .SYNOPSIS
        Generates inline Html style attribute from PScribo list Item style overrides.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Item
    )
    process
    {
        $itemStyleBuilder = New-Object -TypeName System.Text.StringBuilder

        if ($Item.Font)
        {
            $fontList = $List.Font -Join "','"
            [ref] $null = $itemStyleBuilder.AppendFormat(" font-family: '{0}';", $fontList)
        }

        if ($Item.Size -gt 0)
        {
            ## Create culture invariant decimal https://github.com/iainbrighton/PScribo/issues/6
            $invariantItemSize = ConvertTo-InvariantCultureString -Object ($Item.Size / 12) -Format 'f2'
            [ref] $null = $itemStyleBuilder.AppendFormat(' font-size: {0}rem;', $invariantItemSize)
        }

        if ($Item.Bold -eq $true)
        {
            [ref] $null = $itemStyleBuilder.Append(' font-weight: bold;')
        }

        if ($Item.Italic -eq $true)
        {
            [ref] $null = $itemStyleBuilder.Append(' font-style: italic;')
        }

        if ($Item.Underline -eq $true)
        {
            [ref] $null = $itemStyleBuilder.Append(' text-decoration: underline;')
        }

        if (-not [System.String]::IsNullOrEmpty($Item.Color))
        {
            $color = Resolve-PScriboStyleColor -Color $Item.Color
            [ref] $null = $itemStyleBuilder.AppendFormat(' color: #{0};', $color)
        }

        return $itemStyleBuilder.ToString().TrimStart()
    }
}

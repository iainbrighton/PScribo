function Get-HtmlTableStyle
{
<#
    .SYNOPSIS
        Generates html stylesheet style attributes from a PScribo document table style.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        ## PScribo document table style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TableStyle
    )
    process
    {
        $tableStyleBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [ref] $null = $tableStyleBuilder.AppendFormat((Get-HtmlTablePaddingStyle -TableStyle $TableStyle))
        [ref] $null = $tableStyleBuilder.AppendFormat(' border-style: {0};', $TableStyle.BorderStyle.ToLower())

        if ($TableStyle.BorderWidth -gt 0)
        {
            $invariantBorderWidth = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.BorderWidth)
            [ref] $null = $tableStyleBuilder.AppendFormat(' border-width: {0}rem;', $invariantBorderWidth)

            $borderColor = Resolve-PScriboStyleColor -Color $TableStyle.BorderColor
            [ref] $null = $tableStyleBuilder.AppendFormat(' border-color: #{0};', $borderColor.ToLower())
        }

        [ref] $null = $tableStyleBuilder.Append(' border-collapse: collapse;')
        ## <table align="center"> is deprecated in Html5
        if ($TableStyle.Align -eq 'Center')
        {
            [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: auto;')
        }
        elseif ($TableStyle.Align -eq 'Right')
        {
            [ref] $null = $tableStyleBuilder.Append(' margin-left: auto; margin-right: 0;')
        }

        return $tableStyleBuilder.ToString()
    }
}

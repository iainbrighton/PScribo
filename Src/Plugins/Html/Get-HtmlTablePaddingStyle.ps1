function Get-HtmlTablePaddingStyle
{
<#
    .SYNOPSIS
        Generates html stylesheet style padding attributes from a PScribo document table style.
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
        [ref] $null = $tableStyleBuilder.AppendFormat(' padding: {0}rem {1}rem {2}rem {3}rem;',
            (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingTop)),
                (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingRight)),
                    (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingBottom)),
                        (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingLeft)))
        return $tableStyleBuilder.ToString()
    }
}

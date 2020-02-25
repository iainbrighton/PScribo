function GetHtmlPaddedTableStyle
{
<#
    .SYNOPSIS
        Generates padded html stylesheet style attributes from a PScribo table style.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        ## PScribo document table style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TableStyle
    )
    process
    {
        $styleBuilder = New-Object -TypeName System.Text.StringBuilder;

        [ref] $null = $styleBuilder.AppendFormat(' padding: {0}rem {1}rem {2}rem {3}rem;',
            (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingTop)),
                (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingRight)),
                    (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingBottom)),
                        (ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $TableStyle.PaddingLeft)))

        return $styleBuilder.ToString();
    }
}
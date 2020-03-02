function ConvertTo-IndentedString
{
<#
    .SYNOPSIS
        Indents a block of text using the number of tab stops.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [Object[]] $InputObject,

        ## Tab indent
        [Parameter()]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0,

        ## Tab size
        [Parameter()]
        [ValidateRange(0,10)]
        [System.Int32] $TabSize = 4
    )
    process
    {
        $padding = ''.PadRight(($Tabs * $TabSize), ' ')
        ## Use a StringBuilder to write the table line by line (to indent it)
        [System.Text.StringBuilder] $builder = New-Object System.Text.StringBuilder

        foreach ($line in ($InputObject -split '\r\n?|\n'))
        {
            [ref] $null = $builder.Append($padding)
            [ref] $null = $builder.AppendLine($line.TrimEnd()) ## Trim trailing space (#67)
        }
        return $builder.ToString()
    }
}

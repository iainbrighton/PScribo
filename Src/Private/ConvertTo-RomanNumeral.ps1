function ConvertTo-RomanNumeral
{
<#
    .SYNOPSIS
        Converts a decimal number to Roman numerals.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Int32] $Value
    )
    begin
    {
        $conversionTable = [ordered] @{
            1000 = 'M'
            900  = 'CM'
            500  = 'D'
            400  = 'CD'
            100  = 'C'
            90   = 'XC'
            50   = 'L'
            40   = 'XL'
            10   = 'X'
            9    = 'IX'
            5    = 'V'
            4    = 'IV'
            1    = 'I'
        }
    }
    process
    {
        $romanNumeralBuilder = New-Object -TypeName System.Text.StringBuilder
        do
        {
            foreach ($romanNumeral in $conversionTable.GetEnumerator())
            {
                if ($Value -ge $romanNumeral.Key)
                {
                    [ref] $null = $romanNumeralBuilder.Append($romanNumeral.Value)
                    $Value -= $romanNumeral.Key
                    break
                }
            }

        }
        until ($Value -eq 0)
        return $romanNumeralBuilder.ToString()
    }
}

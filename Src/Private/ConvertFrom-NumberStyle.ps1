function ConvertFrom-NumberStyle
{
<#
    .SYNOPSIS
        Converts a number to its string representation, based upon the number style
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Int32] $Value,

        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $NumberStyle
    )
    process
    {
        switch ($NumberStyle.Format)
        {
            'Number'
            {
                $numberString = $Value.ToString()
            }
            'Letter'
            {
                $numberString = [System.Char] ($Value + 64)
            }
            'Roman'
            {
                $numberString = ConvertTo-RomanNumeral -Value $Value
            }
            'Custom'
            {
                $numberCount = 0
                $customStringChars = $numberStyle.Custom.ToCharArray()
                $customStringLength = $numberStyle.Custom.Length
                for ($n = $customStringLength - 1; $n -ge 0; $n--)
                {
                    if ($customStringChars[$n] -eq '%')
                    {
                        $numberCount++
                    }
                }

                $searchString = ''.PadRight($numberCount, '%')
                $replacementString = $Value.ToString().PadLeft($numberCount, '0')
                return $numberStyle.Custom.Replace($searchString, $replacementString)
            }
        }

        $numberString = '{0}{1}' -f $numberString, $NumberStyle.Suffix
        if ($NumberStyle.Uppercase)
        {
            return $numberString.ToUpper()
        }
        else
        {
            return $numberString.ToLower()
        }
    }
}

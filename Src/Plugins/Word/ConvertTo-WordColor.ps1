function ConvertTo-WordColor
{
<#
    .SYNOPSIS
        Converts an HTML color to RRGGBB value as Word does not support short Html color codes
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.String] $Color
    )
    process
    {
        $Color = $Color.TrimStart('#')
        if ($Color.Length -eq 3)
        {
            $Color = '{0}{0}{1}{1}{2}{2}' -f $Color[0], $Color[1], $Color[2]
        }

        return $Color.ToUpper()
    }
}

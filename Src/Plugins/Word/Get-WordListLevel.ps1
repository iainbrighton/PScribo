function Get-WordListLevel
{
<#
    .SYNOPSIS
        Process (nested) lists and build index of number/bullet styles.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Collections.Hashtable] $Levels = @{ }
    )
    process
    {
        foreach ($item in $List.Items)
        {
            if ($item.Type -eq 'PScribo.Item')
            {
                $level = $item.Level -1
                if (-not $Levels.ContainsKey($level))
                {
                    $Levels[$level] = @{
                        IsNumbered  = $List.IsNumbered
                        NumberStyle = $List.NumberStyle
                        BulletStyle = $List.BulletStyle
                    }
                }
            }
            elseif ($item.Type -eq 'PScribo.List')
            {
                $Levels = Get-WordListLevel -List $item -Levels $Levels
            }
        }
        return $Levels
    }
}

function Get-WordParagraphRun
{
<#
    .SYNOPSIS
        Returns an array of strings, split on tokens (PageNumber/TotalPages).
#>
    [CmdletBinding()]
    [OutputType([System.String[]])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [System.String] $Text,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoSpace
    )
    process
    {
        $pageNumberMatch = '<!#\s*PageNumber\s*#!>'
        $totalPagesMatch = '<!#\s*TotalPages\s*#!>'
        $runs = New-Object -TypeName 'System.Collections.ArrayList'

        if (-not $NoSpace)
        {
            $Text = '{0} ' -f $Text
        }

        $pageNumbers = $Text -isplit $pageNumberMatch
        for ($pn = 0; $pn -lt $pageNumbers.Count; $pn++)
        {
            $totalPages = $pageNumbers[$pn] -isplit $totalPagesMatch
            for ($tp = 0; $tp -lt $totalPages.Count; $tp++)
            {
                if (-not [System.String]::IsNullOrWhitespace($totalPages[$tp]))
                {
                    $null = $runs.Add($totalPages[$tp])
                }
                if ($tp -lt ($totalPages.Count -1))
                {
                    $null = $runs.Add('<!#TOTALPAGES#!>')
                }
            }

            if ($pn -lt ($pageNumbers.Count -1))
            {
                $null = $runs.Add('<!#PAGENUMBER#!>')
            }
        }

        return $runs.ToArray()
    }
}

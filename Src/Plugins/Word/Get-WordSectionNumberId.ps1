function Get-WordSectionNumberId
{
<#
    .SYNOPSIS
        Returns the Word numbering id of the section heading list.

    .NOTES
        Bullet/numbered lists occupy the ids 1 to n, so the section headings use the next one.
#>
    [CmdletBinding()]
    [OutputType([System.Int32])]
    param
    (
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList] $Lists
    )
    process
    {
        return $Lists.Count + 1
    }
}

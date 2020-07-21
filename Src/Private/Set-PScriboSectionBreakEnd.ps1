function Set-PScriboSectionBreakEnd
{
<#
    .SYNOPSIS
        Sets the IsSectionBreakEnd on the last (nested) paragraph/subsection (required by Word plugin)
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Section
    )
    process
    {
        $Section.Sections |
            Where-Object { $_.Type -in 'PScribo.Section','PScribo.Paragraph','PScribo.Table' } |
            Select-Object -Last 1 | ForEach-Object {
                    if ($PSItem.Type -in 'PScribo.Paragraph','PScribo.Table')
                    {
                        $PSItem.IsSectionBreakEnd = $true
                    }
                    else
                    {
                        Set-PScriboSectionBreakEnd -Section $PSItem
                    }
                }
    }
}

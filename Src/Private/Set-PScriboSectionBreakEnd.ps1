function Set-PScriboSectionBreakEnd
{
<#
    .SYNOPSIS
        Sets the IsSectionBreak end on the last (nested) paragraph/subsection (required by Word plugin)
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
            Where-Object { $_.Type -in 'PScribo.Section','PScribo.Paragraph' } |
                Select-Object -Last 1 | ForEach-Object {
                    if ($PSItem.Type -eq 'PScribo.Paragraph')
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

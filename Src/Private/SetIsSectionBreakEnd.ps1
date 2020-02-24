function SetIsSectionBreakEnd
{
<#
    Sets the IsSectionBreak end on the last (nested) paragraph/subsection (required by Word plugin)
#>
    [CmdletBinding()]
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
                        SetIsSectionBreakEnd -Section $PSItem
                    }
                }
    }
}

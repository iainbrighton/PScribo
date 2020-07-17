function Get-HtmlTableColGroup
{
<#
    .SYNOPSIS
        Generates Html <colgroup> tags based on table column widths
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        $colGroupBuilder = New-Object -TypeName 'System.Text.StringBuilder'

        if ($Table.ColumnWidths)
        {
            [ref] $null = $colGroupBuilder.Append('<colgroup>')
            foreach ($columnWidth in $Table.ColumnWidths)
            {
                if ($null -eq $columnWidth)
                {
                    [ref] $null = $colGroupBuilder.Append('<col />')
                }
                else
                {
                    [ref] $null = $colGroupBuilder.AppendFormat('<col style="max-width:{0}%; min-width:{0}%; width:{0}%" />', $columnWidth)
                }
            }
            [ref] $null = $colGroupBuilder.AppendLine('</colgroup>')
        }

        return $colGroupBuilder.ToString()
    }
}

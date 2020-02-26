function GetHtmlTable
{
<#
    .SYNOPSIS
        Generates html <table> from a PScribo.Table object.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    begin
    {
        $formattedTable = ConvertTo-PScriboPreformattedTable -Table $Table
    }
    process
    {
        $tableBuilder = New-Object -TypeName System.Text.StringBuilder
        [ref] $null = $tableBuilder.Append((GetHtmlTableDiv -Table $Table))
        [ref] $null = $tableBuilder.Append((GetHtmlTableColGroup -Table $Table))

        ## Table headers
        $startRow = 0
        if ($formattedTable.HasHeaderRow)
        {
            [ref] $null = $tableBuilder.Append('<thead><tr>')
            foreach ($cell in $formattedTable.Rows[0].Cells)
            {
                [ref] $null = $tableBuilder.AppendFormat('<th>{0}</th>', $cell.Content)
            }
            [ref] $null = $tableBuilder.Append('</tr></thead>')
            $startRow += 1
        }

        ## Table rows
        [ref] $null = $tableBuilder.AppendLine('<tbody>')
        for ($r = $startRow; $r -lt $formattedTable.Rows.Count; $r++)
        {
            $row = $formattedTable.Rows[$r]

            if ($row.IsStyleInherited)
            {
                [ref] $null = $tableBuilder.Append('<tr>')
            }
            else
            {
                [ref] $null = $tableBuilder.AppendFormat('<tr style="{0}">', $row.Style)
            }

            for ($c = 0; $c -lt $row.Cells.Count; $c++)
            {
                $cell = $row.Cells[$c]

                if ([System.String]::IsNullOrEmpty($cell.Content))
                {
                    $encodedHtmlContent = '&nbsp;' # &nbsp; is already encoded (#72)
                }
                else
                {
                    $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($cell.Content)
                    $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')
                }

                if ($formattedTable.HasHeaderColumn -and ($c -eq 0))
                {
                    ## Display first column with header styling
                    [ref] $null = $tableBuilder.AppendFormat('<th>{0}</th>', $encodedHtmlContent)
                }
                else
                {
                    if ($cell.IsStyleInherited)
                    {
                        [ref] $null = $tableBuilder.AppendFormat('<td>{0}</td>', $encodedHtmlContent)
                    }
                    else
                    {
                        $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$cell.Style]).Trim()
                        [ref] $null = $tableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent)
                    }
                }
            }
            [ref] $null = $tableBuilder.AppendLine('</tr>')
        }
        [ref] $null = $tableBuilder.AppendLine('</tbody></table></div>')
        return $tableBuilder.ToString()
    }
}

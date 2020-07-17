function Get-HtmlTable
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
        $formattedTables = ConvertTo-PScriboPreformattedTable -Table $Table
    }
    process
    {
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style
        $tableBuilder = New-Object -TypeName System.Text.StringBuilder
        foreach ($formattedTable in $formattedTables)
        {
            if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Above'))
            {
                [ref] $null = $tableBuilder.Append((Get-HtmlTableCaption -Table $Table))
            }

            [ref] $null = $tableBuilder.Append((Get-HtmlTableDiv -Table $Table))
            [ref] $null = $tableBuilder.Append((Get-HtmlTableColGroup -Table $Table))

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
                    [ref] $null = $tableBuilder.AppendFormat('<tr style="{0}">', (Get-HtmlStyle -Style $Document.Styles[$row.Style]).Trim())
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
                        $cellContent = Resolve-PScriboToken -InputObject $cell.Content
                        $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($cellContent)
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
                            $propertyStyleHtml = (Get-HtmlStyle -Style $Document.Styles[$cell.Style]).Trim()
                            [ref] $null = $tableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent)
                        }
                    }
                }
                [ref] $null = $tableBuilder.AppendLine('</tr>')
            }

            [ref] $null = $tableBuilder.AppendLine('</tbody></table>')

            if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Below'))
            {
                [ref] $null = $tableBuilder.Append((Get-HtmlTableCaption -Table $Table))
            }

            ## Add a space between each table to mirror Word output rendering
            [ref] $null = $tableBuilder.AppendLine('<br /></div>')
        }
        return $tableBuilder.ToString()
    }
}

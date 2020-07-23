function Out-MarkdownTable
{
<#
    .SYNOPSIS
        Output formatted table from PScribo.Table object.
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
        if ($script:currentPScriboObject -eq 'PScribo.Paragraph')
        {
            [ref] $null = $tableBuilder.AppendLine()
        }

        if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Above'))
        {
            $tableCaption = Get-MarkdownTableCaption -Table $Table
            [ref] $null = $tableBuilder.AppendLine($tableCaption)
        }

        foreach ($formattedTable in $formattedTables)
        {
            ## Table headers
            $headerBuilder = New-Object -TypeName System.Text.StringBuilder
            foreach ($cell in $formattedTable.Rows[0].Cells)
            {
                $cellContent = $cell.Content
                [ref] $null = $tableBuilder.AppendFormat('| {0} ', $cellContent)

                ## Markdown table headers require at least one hyphen, i.e. :-
                $cellContentLength = $cellContent.Length
                if ($cellContentLength -le 1)
                {
                    $cellContentLength = 2
                }
                $headerMarker = ':'.PadRight($cellContentLength, '-')
                [ref] $null = $headerBuilder.AppendFormat('| {0} ', $headerMarker)
            }
            [ref] $null = $tableBuilder.AppendLine().AppendLine($headerBuilder.ToString())

            ## Table rows
            for ($r = 1; $r -lt $formattedTable.Rows.Count; $r++)
            {
                $row = $formattedTable.Rows[$r]
                $rowStyle = ''
                if (-not $row.IsStyleInherited)
                {
                    $rowStyle = Get-MarkdownInlineStyle -Style $row.Style
                }

                for ($c = 0; $c -lt $row.Cells.Count; $c++)
                {
                    $cell = $row.Cells[$c]
                    if ($cell.IsStyleInherited)
                    {
                        $cellStyle = $rowStyle
                    }
                    else
                    {
                        $cellStyle = Get-MarkdownInlineStyle -Style $cell.Style
                    }
                    $cellContent = $cell.Content.Replace('|', '\|')
                    [ref] $null = $tableBuilder.AppendFormat('| {0}{1}{0} ', $cellStyle, $cellContent)
                }

                [ref] $null = $tableBuilder.AppendLine()
            }

            ## Add a space between each table to mirror Word output rendering
            [ref] $null = $tableBuilder.AppendLine()
        }

        if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Below'))
        {
            $tableCaption = Get-MarkdownTableCaption -Table $Table
            [ref] $null = $tableBuilder.AppendLine($tableCaption)
        }

        $script:currentPScriboObject = 'PScribo.Table'
        return $tableBuilder.ToString()
    }
}

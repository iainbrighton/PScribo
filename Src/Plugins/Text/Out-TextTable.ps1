function Out-TextTable
{
<#
    .SYNOPSIS
        Output formatted text table.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Table
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboTextOption
        }
    }
    process
    {
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style
        $tableBuilder = New-Object -TypeName System.Text.StringBuilder
        $tableRenderWidth = $options.TextWidth - ($Table.Tabs * 4)

        ## We need to replace page numbers before outputting the table
        foreach ($row in $Table.Rows)
        {
            foreach ($property in $row.PSObject.Properties)
            {
                if ($property.Value -is [System.Array])
                {
                    $property.Value = [System.String]::Join(' ', $property.Value)
                }
                $property.Value = Resolve-PScriboToken -InputObject $property.Value
            }
        }

        ## We've got to render the table first to determine how wide it is
        ## before we can justify it
        if ($Table.IsKeyedList)
        {
            ## Create new objects with headings as properties
            $tableText = (ConvertTo-PSObjectKeyedListTable -Table $Table |
                            Select-Object -Property * -ExcludeProperty '*__Style' |
                            Format-Table -Wrap -AutoSize |
                                Out-String -Width $tableRenderWidth).Trim([System.Environment]::NewLine)
        }
        elseif ($Table.IsList)
        {
            $tableText = ($Table.Rows |
                Select-Object -Property * -ExcludeProperty '*__Style' |
                    Format-List | Out-String -Width $tableRenderWidth).Trim([System.Environment]::NewLine)
        }
        else
        {
            ## Don't trim tabs for table headers
            ## Tables set to AutoSize as otherwise rendering is different between PoSh v4 and v5
            $tableText = ($Table.Rows |
                            Select-Object -Property * -ExcludeProperty '*__Style' |
                                Format-Table -Wrap -AutoSize |
                                    Out-String -Width $tableRenderWidth).Trim([System.Environment]::NewLine)
        }

        if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Above'))
        {
            $justifiedCaption = Get-TextTableCaption -Table $Table
            [ref] $null = $tableBuilder.AppendLine($justifiedCaption)
            [ref] $null = $tableBuilder.AppendLine()
        }

        ## We don't want to wrap table contents so just justify it
        $convertToJustifiedStringParams = @{
            InputObject = $tableText
            Width = $tableRenderWidth
            Align = $tableStyle.Align
        }
        $justifiedTableText = ConvertTo-JustifiedString @convertToJustifiedStringParams

        [ref] $null = $tableBuilder.Append($justifiedTableText)

        if ($Table.HasCaption -and ($tableStyle.CaptionLocation -eq 'Below'))
        {
            $justifiedCaption = Get-TextTableCaption -Table $Table
            [ref] $null = $tableBuilder.AppendLine()
            [ref] $null = $tableBuilder.AppendLine()
            [ref] $null = $tableBuilder.Append($justifiedCaption)
        }

        $convertToIndentedStringParams = @{
            InputObject = $tableBuilder.ToString()
            Tabs        = $Table.Tabs
        }

        return (ConvertTo-IndentedString @convertToIndentedStringParams)
    }
}

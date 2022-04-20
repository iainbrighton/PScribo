function Get-HtmlTableDiv
{
<#
    .SYNOPSIS
        Generates Html <div style=..><table style=..> tags based upon table width, columns and indentation
    .NOTES
        A <div> is required to ensure that the table stays within the "page" boundaries/margins.
#>
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        $divBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [ref] $null = $divBuilder.Append('<div style="word-break: break-word; overflow-wrap: anywhere; ')

        if ($Table.Tabs -gt 0)
        {
            $invariantMarginLeft = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Table.Tabs))
            [ref] $null = $divBuilder.AppendFormat('margin-left: {0}rem; ' -f $invariantMarginLeft)
        }
        ## Ensure we close the <div style=" "> tag
        [ref] $null = $divBuilder.AppendFormat('"><table class="{0}"', $Table.Style.ToLower())

        $styleElements = @()
        if ($Table.Width -gt 0)
        {
            $styleElements += 'width:{0}%;' -f $Table.Width
        }

        if ($Table.ColumnWidths)
        {
            $styleElements += 'table-layout: fixed;'
            #$styleElements += 'word-break: break-word;' # 'word-wrap: break-word;' or 'overflow-wrap: break-word;'?
        }

        if ($styleElements.Count -gt 0)
        {
            [ref] $null = $divBuilder.AppendFormat(' style="{0}">', [System.String]::Join(' ', $styleElements))
        }
        else
        {
            [ref] $null = $divBuilder.Append('>')
        }

        return $divBuilder.ToString()
    }
}

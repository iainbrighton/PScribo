function GetHtmlTableDiv
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
        $divBuilder = New-Object -TypeName 'System.Text.StringBuilder';
        if ($Table.Tabs -gt 0)
        {
            $invariantMarginLeft = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Table.Tabs));
            [ref] $null = $divBuilder.AppendFormat('<div style="margin-left: {0}rem;">' -f $invariantMarginLeft);
        }
        else
        {
            [ref] $null = $divBuilder.Append('<div>');
        }

        [ref] $null = $divBuilder.AppendFormat('<table class="{0}"', $Table.Style.ToLower());

        $styleElements = @();
        if ($Table.Width -gt 0)
        {
            $styleElements += 'width:{0}%;' -f $Table.Width;
        }
        if ($Table.ColumnWidths)
        {
            $styleElements += 'table-layout: fixed;';
            $styleElements += 'word-break: break-word;'
        }
        if ($styleElements.Count -gt 0)
        {
            [ref] $null = $divBuilder.AppendFormat(' style="{0}">', [System.String]::Join(' ', $styleElements));
        }
        else
        {
            [ref] $null = $divBuilder.Append('>');
        }
        return $divBuilder.ToString();
    }
}

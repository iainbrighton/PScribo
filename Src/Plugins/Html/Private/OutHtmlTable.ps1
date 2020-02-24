function OutHtmlTable
{
<#
    .SYNOPSIS
        Output formatted Html <table> from PScribo.Table object.
    .NOTES
        One table is output per table row with the -List parameter.
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
        [System.Text.StringBuilder] $tableBuilder = New-Object -TypeName 'System.Text.StringBuilder';
        if ($Table.IsList)
        {
            ## Create a table for each row
            for ($r = 0; $r -lt $Table.Rows.Count; $r++)
            {
                $row = $Table.Rows[$r];
                if ($r -gt 0)
                {
                    ## Add a space between each table to mirror Word output rendering
                    [ref] $null = $tableBuilder.AppendLine('<p />');
                }
                [ref] $null = $tableBuilder.Append((GetHtmlTableList -Table $Table -Row $row));
            } #end foreach row
        }
        else
        {
            [ref] $null = $tableBuilder.Append((GetHtmlTable -Table $Table));
        }
        return $tableBuilder.ToString();
    }
}

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
    process
    {
        $standardTableBuilder = New-Object -TypeName System.Text.StringBuilder
        [ref] $null = $standardTableBuilder.Append((GetHtmlTableDiv -Table $Table))
        [ref] $null = $standardTableBuilder.Append((GetHtmlTableColGroup -Table $Table))

        ## Table headers
        [ref] $null = $standardTableBuilder.Append('<thead><tr>')
        for ($i = 0; $i -lt $Table.Columns.Count; $i++)
        {
            [ref] $null = $standardTableBuilder.AppendFormat('<th>{0}</th>', $Table.Columns[$i])
        }
        [ref] $null = $standardTableBuilder.Append('</tr></thead>')

        ## Table body
        [ref] $null = $standardTableBuilder.AppendLine('<tbody>')
        foreach ($row in $Table.Rows)
        {
            [ref] $null = $standardTableBuilder.Append('<tr>')
            foreach ($propertyName in $Table.Columns)
            {
                $propertyStyle = '{0}__Style' -f $propertyName

                $rowPropertyName = $row.$propertyName; ## Core
                if ([System.String]::IsNullOrEmpty($rowPropertyName))
                {
                    $encodedHtmlContent = '&nbsp;' # &nbsp; is already encoded (#72)
                }
                else
                {
                    $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($rowPropertyName.ToString())
                }
                $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')

                if ($row.PSObject.Properties[$propertyStyle])
                {
                    ## Cell styles override row styles
                    $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.$propertyStyle]).Trim()
                    [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $propertyStyleHtml, $encodedHtmlContent)
                }
                elseif (($row.PSObject.Properties['__Style']) -and (-not [System.String]::IsNullOrEmpty($row.__Style)))
                {
                    ## We have a row style
                    $rowStyleHtml = (GetHtmlStyle -Style $Document.Styles[$row.__Style]).Trim()
                    [ref] $null = $standardTableBuilder.AppendFormat('<td style="{0}">{1}</td>', $rowStyleHtml, $encodedHtmlContent)
                }
                else
                {
                    if ($null -ne $row.$propertyName)
                    {
                        ## Check that the property has a value
                        [ref] $null = $standardTableBuilder.AppendFormat('<td>{0}</td>', $encodedHtmlContent)
                    }
                    else
                    {
                        [ref] $null = $standardTableBuilder.Append('<td>&nbsp;</td>')
                    }
                } #end if $row.PropertyStyle
            } #end foreach property
            [ref] $null = $standardTableBuilder.AppendLine('</tr>')
        } #end foreach row
        [ref] $null = $standardTableBuilder.AppendLine('</tbody></table></div>')
        return $standardTableBuilder.ToString()
    }
}

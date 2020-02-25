function GetHtmlTableList
{
<#
    .SYNOPSIS
        Generates list html <table> from a PScribo.Table row object.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table,

        [Parameter(Mandatory)]
        [System.Management.Automation.PSObject] $Row
    )
    process
    {
        $listTableBuilder = New-Object -TypeName System.Text.StringBuilder
        [ref] $null = $listTableBuilder.Append((GetHtmlTableDiv -Table $Table))
        [ref] $null = $listTableBuilder.Append((GetHtmlTableColGroup -Table $Table))
        [ref] $null = $listTableBuilder.Append('<tbody>')

        for ($i = 0; $i -lt $Table.Columns.Count; $i++)
        {
            $propertyName = $Table.Columns[$i]
            $rowPropertyName = $Row.$propertyName ## Core
            [ref] $null = $listTableBuilder.AppendFormat('<tr><th>{0}</th>', $propertyName)
            $propertyStyle = '{0}__Style' -f $propertyName

            if ($row.PSObject.Properties[$propertyStyle])
            {
                $propertyStyleHtml = (GetHtmlStyle -Style $Document.Styles[$Row.$propertyStyle])
                if ([System.String]::IsNullOrEmpty($rowPropertyName))
                {

                    [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">&nbsp;</td></tr>', $propertyStyleHtml)
                }
                else
                {
                    $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString())
                    $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')
                    [ref] $null = $listTableBuilder.AppendFormat('<td style="{0}">{1}</td></tr>', $propertyStyleHtml, $encodedHtmlContent)
                }
            }
            else
            {
                if ([System.String]::IsNullOrEmpty($rowPropertyName))
                {
                    [ref] $null = $listTableBuilder.Append('<td>&nbsp;</td></tr>')
                }
                else
                {
                    $encodedHtmlContent = [System.Net.WebUtility]::HtmlEncode($row.$propertyName.ToString())
                    $encodedHtmlContent = $encodedHtmlContent.Replace([System.Environment]::NewLine, '<br />')
                    [ref] $null = $listTableBuilder.AppendFormat('<td>{0}</td></tr>', $encodedHtmlContent)
                }
            }
        } #end for each property
        [ref] $null = $listTableBuilder.AppendLine('</tbody></table></div>')
        return $listTableBuilder.ToString()
    }
}

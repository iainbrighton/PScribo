function Out-HtmlTOC
{
<#
    .SYNOPSIS
        Generates Html table of contents.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TOC
    )
    process
    {
        $tocBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [ref] $null = $tocBuilder.AppendFormat('<h1 class="{0}">{1}</h1>', $TOC.ClassId, $TOC.Name)
        #[ref] $null = $tocBuilder.AppendLine('<table style="width: 100%;">')
        [ref] $null = $tocBuilder.AppendLine('<table>')
        foreach ($tocEntry in $Document.TOC)
        {
            $sectionNumberIndent = '&nbsp;&nbsp;&nbsp;' * $tocEntry.Level
            if ($Document.Options['EnableSectionNumbering'])
            {
                [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}</td><td>{1}<a href="#{2}" style="text-decoration: none;">{3}</a></td></tr>', $tocEntry.Number, $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine()
            }
            else
            {
                [ref] $null = $tocBuilder.AppendFormat('<tr><td>{0}<a href="#{1}" style="text-decoration: none;">{2}</a></td></tr>', $sectionNumberIndent, $tocEntry.Id, $tocEntry.Name).AppendLine()
            }
        }
        [ref] $null = $tocBuilder.AppendLine('</table>')
        return $tocBuilder.ToString()
    }
}

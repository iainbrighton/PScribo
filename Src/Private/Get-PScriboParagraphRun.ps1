function Get-PScriboParagraphRun
{
<#
    .SYNOPSIS
        Converts a string of text into multiple strings to create Word
        paragraph runs for PageNumber and TotalPages field replacements.
#>
    [CmdletBinding()]
    param
    (
        ## Paragraph text to convert into Word paragraph runs for field replacement
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.String] $Text,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlDocument] $XmlDocument,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlElement] $XmlElement
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        $pageNumberMatch = '<!#\s*PageNumber\s*#!>'
        $totalPagesMatch = '<!#\s*TotalPages\s*#!>'
        $paragraphRuns = New-Object -TypeName 'System.Collections.ArrayList'

        $pageNumbers = $Text -split $pageNumberMatch
        for ($pn = 0; $pn -lt $pageNumbers.Count; $pn++)
        {
            $totalPages = $pageNumbers[$pn] -split $totalPagesMatch
            for ($tp = 0; $tp -lt $totalPages.Count; $tp++)
            {
                $null = $paragraphRuns.Add($totalPages[$tp])

                if ($tp -lt ($totalPages.Count -1))
                {
                    $null = $paragraphRuns.Add('<!#TOTALPAGES#!>')
                }
            }

            if ($pn -lt ($pageNumbers.Count -1))
            {
                $null = $paragraphRuns.Add('<!#PAGENUMBER#!>')
            }
        }

        foreach ($run in $paragraphRuns)
        {
            if ($run -match '<!#(TOTALPAGES|PAGENUMBER)#!>')
            {
                $r1 = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $fldChar1 = $r1.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlns, 'begin')

                $r2 = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $instrText = $r2.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlns))
                [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')

                if ($run -match '<!#PAGENUMBER#!>')
                {
                    [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' PAGE   \* MERGEFORMAT '))
                }
                elseif ($run -match '<!#TOTALPAGES#!>')
                {
                    [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' NUMPAGES   \* MERGEFORMAT '))
                }

                $r3 = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $fldChar2 = $r3.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlns, 'separate')

                $r4 = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $t2 = $r4.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
                [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode('1'))

                $r5 = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlns, 'end')
            }
            else
            {
                $r = $XmlElement.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
                [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($run))
            }
        }
    }
}

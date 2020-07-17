function Out-WordParagraph
{
<#
    .SYNOPSIS
        Output formatted Word paragraph and run(s).
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Paragraph,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlns);
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns));

        if ($Paragraph.Tabs -gt 0)
        {
            $ind = $pPr.AppendChild($XmlDocument.CreateElement('w', 'ind', $xmlns))
            [ref] $null = $ind.SetAttribute('left', $xmlns, (720 * $Paragraph.Tabs))
        }
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlns))
            [ref] $null = $pStyle.SetAttribute('val', $xmlns, $Paragraph.Style)
        }

        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('before', $xmlns, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlns, 0)

        if ($Paragraph.IsSectionBreakEnd)
        {
            $paragraphPrParams = @{
                PageHeight       = $Document.Options['PageHeight']
                PageWidth        = $Document.Options['PageWidth']
                PageMarginTop    = $Document.Options['MarginTop'];
                PageMarginBottom = $Document.Options['MarginBottom'];
                PageMarginLeft   = $Document.Options['MarginLeft'];
                PageMarginRight  = $Document.Options['MarginRight'];
                Orientation      = $Paragraph.Orientation;
            }
            [ref] $null = $pPr.AppendChild((Get-WordSectionPr @paragraphPrParams -XmlDocument $xmlDocument))
        }

        foreach ($paragraphRun in $Paragraph.Sections)
        {
            $rPr = Get-WordParagraphRunPr -ParagraphRun $paragraphRun -XmlDocument $XmlDocument
            $noSpace = ($paragraphRun.IsParagraphRunEnd -eq $true) -or ($paragraphRun.NoSpace -eq $true)
            $runs = Get-WordParagraphRun -Text $paragraphRun.Text -NoSpace:$noSpace

            foreach ($run in $runs)
            {
                if (-not [System.String]::IsNullOrEmpty($run))
                {
                    if ($run -imatch '<!#(TOTALPAGES|PAGENUMBER)#!>')
                    {
                        $r1 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        $fldChar1 = $r1.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                        [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlns, 'begin')

                        $r2 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        [ref] $null = $r2.AppendChild($rPr)
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

                        $r3 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        $fldChar2 = $r3.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                        [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlns, 'separate')

                        $r4 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        [ref] $null = $r4.AppendChild($rPr)
                        $t2 = $r4.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
                        [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode('1'))

                        $r5 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
                        [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlns, 'end')
                    }
                    else
                    {
                        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                        [ref] $null = $r.AppendChild($rPr)

                        ## Create a separate text block for each line/break
                        $lines = $run -split '\r\n?|\n'
                        for ($l = 0; $l -lt $lines.Count; $l++)
                        {
                            $line = $lines[$l]
                            $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
                            if ($line -ne $line.Trim())
                            {
                                ## Only preserve space if there is a preceeding or trailing space
                                [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
                            }
                            [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($line))

                            if ($l -lt ($lines.Count - 1))
                            {
                                ## Don't add a line break to the last line/break
                                [ref] $null = $r.AppendChild($XmlDocument.CreateElement('w', 'br', $xmlns))
                            }
                        }
                    }
                }
            }
        }

        return $p
    }
}

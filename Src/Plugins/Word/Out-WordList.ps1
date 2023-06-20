function Out-WordList
{
<#
    .SYNOPSIS
        Output formatted Word list.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlElement] $Element,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlDocument] $XmlDocument,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $NumberId,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NotRoot
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

        foreach ($item in $List.Items)
        {
            if ($item.Type -eq 'PScribo.Item')
            {
                $p = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
                $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))

                if ($List.HasStyle)
                {
                    $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlns))
                    [ref] $null = $pStyle.SetAttribute('val', $xmlns, $List.Style)
                }

                $numPr = $pPr.AppendChild($XmlDocument.CreateElement('w', 'numPr', $xmlns))
                $ilvl = $numPr.AppendChild($XmlDocument.CreateElement('w', 'ilvl', $xmlns))
                [ref] $null = $ilvl.SetAttribute('val', $xmlns, $item.Level -1)
                $numId = $numPr.AppendChild($XmlDocument.CreateElement('w', 'numId', $xmlns))
                [ref] $null = $numId.SetAttribute('val', $xmlns, $NumberId)

                $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
                $rPr = Get-WordParagraphRunPr -ParagraphRun $item -XmlDocument $XmlDocument
                [ref] $null = $r.AppendChild($rPr)

                $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
                [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($item.Text))
            }
            elseif ($item.Type -eq 'PScribo.List')
            {
                Out-WordList -List $item -Element $Element -XmlDocument $XmlDocument -NumberId $NumberId -NotRoot
            }
        }

        ## Append a blank line after each list
        if (-not $NotRoot)
        {
            $p = $Element.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        }
    }
}

function Get-WordTableCaption
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style
        $caption = ' {0}' -f $Table.Caption

        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $p = $XmlDocument.CreateElement('w', 'p', $xmlns)
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        $pStyle = $pPr.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlns))
        [ref] $null = $pStyle.SetAttribute('val', $xmlns, $tableStyle.CaptionStyle)

        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $t = $r.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
        [ref] $null = $t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        $paddedTableCaption = '{0} ' -f $tableStyle.CaptionPrefix
        [ref] $null = $t.AppendChild($XmlDocument.CreateTextNode($paddedTableCaption))

        $fldSimple = $p.AppendChild($XmlDocument.CreateElement('w', 'fldSimple', $xmlns))
        [ref] $null = $fldSimple.SetAttribute('instr', $xmlns, ' SEQ Table \* ARABIC ')
        $r2 = $fldSimple.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $rPr2 = $r2.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlns))
        [ref] $null = $rPr2.AppendChild($XmlDocument.CreateElement('w', 'noProof', $xmlns))
        $t2 = $r2.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
        [ref] $null = $t2.AppendChild($XmlDocument.CreateTextNode($Table.CaptionNumber))

        $r3 = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $rPr3 = $r3.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlns))
        [ref] $null = $rPr3.AppendChild($XmlDocument.CreateElement('w', 'noProof', $xmlns))
        $t3 = $r3.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
        [ref] $null = $t3.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        [ref] $null = $t3.AppendChild($XmlDocument.CreateTextNode($caption))

        return $p
    }
}

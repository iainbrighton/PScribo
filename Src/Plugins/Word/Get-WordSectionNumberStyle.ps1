function Get-WordSectionNumberStyle
{
<#
    .SYNOPSIS
        Generates the Word multilevel list definition that numbers section headings.

    .NOTES
        Renders the same '1.2.3 Name' format as the literal section numbering.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory)]
        [System.Int32] $AbstractNumberId,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $abstractNum = $XmlDocument.CreateElement('w', 'abstractNum', $xmlns)
        [ref] $null = $abstractNum.SetAttribute('abstractNumId', $xmlns, $AbstractNumberId)
        $multiLevelType = $abstractNum.AppendChild($XmlDocument.CreateElement('w', 'multiLevelType', $xmlns))
        [ref] $null = $multiLevelType.SetAttribute('val', $xmlns, 'multilevel')

        ## Word supports up to nine list levels
        foreach ($level in 0..8)
        {
            $lvl = $abstractNum.AppendChild($XmlDocument.CreateElement('w', 'lvl', $xmlns))
            [ref] $null = $lvl.SetAttribute('ilvl', $xmlns, $level)
            $start = $lvl.AppendChild($XmlDocument.CreateElement('w', 'start', $xmlns))
            [ref] $null = $start.SetAttribute('val', $xmlns, 1)
            $numFmt = $lvl.AppendChild($XmlDocument.CreateElement('w', 'numFmt', $xmlns))
            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'decimal')
            ## Word defaults to a tab, which indents the heading text beyond the literal section numbering
            $suff = $lvl.AppendChild($XmlDocument.CreateElement('w', 'suff', $xmlns))
            [ref] $null = $suff.SetAttribute('val', $xmlns, 'space')
            $lvlText = $lvl.AppendChild($XmlDocument.CreateElement('w', 'lvlText', $xmlns))
            $levelNumbers = 1..($level + 1) | ForEach-Object { '%{0}' -f $_ }
            [ref] $null = $lvlText.SetAttribute('val', $xmlns, ($levelNumbers -join '.'))
            $lvlJc = $lvl.AppendChild($XmlDocument.CreateElement('w', 'lvlJc', $xmlns))
            [ref] $null = $lvlJc.SetAttribute('val', $xmlns, 'left')
        }

        return $abstractNum
    }
}

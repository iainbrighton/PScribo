function Out-WordTOC
{
<#
    .SYNOPSIS
        Output formatted Word table of contents.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TOC,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlns)
        $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr', $xmlns))
        $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', 'docPartObj', $xmlns))
        $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartGallery', $xmlns))
        [ref] $null = $docObjectGallery.SetAttribute('val', $xmlns, 'Table of Contents')
        [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartUnique', $xmlns))
        [ref] $null = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdEndPr', $xmlns))

        $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdContent', $xmlns))
        $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlns))
        [ref] $null = $pStyle1.SetAttribute('val', $xmlns, 'TOC')
        $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlns))
        [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name))

        $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs', $xmlns))
        $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', $xmlns))
        [ref] $null = $tab2.SetAttribute('val', $xmlns, 'right')
        [ref] $null = $tab2.SetAttribute('leader', $xmlns, 'dot')
        [ref] $null = $tab2.SetAttribute('pos', $xmlns, '9016')

        $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
        [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlns, 'begin')

        $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlns))
        [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' TOC \h \z \u '))

        $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
        [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlns, 'separate')

        $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlns))
        $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlns))
        $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlns))
        [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlns, 'end')

        return $sdt
    }
}

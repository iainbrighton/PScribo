function OutWordTOC
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
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $sdt = $XmlDocument.CreateElement('w', 'sdt', $xmlnsMain)
        $sdtPr = $sdt.AppendChild($XmlDocument.CreateElement('w', 'sdtPr', $xmlnsMain))
        $docPartObj = $sdtPr.AppendChild($XmlDocument.CreateElement('w', 'docPartObj', $xmlnsMain))
        $docObjectGallery = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartGallery', $xmlnsMain))
        [ref] $null = $docObjectGallery.SetAttribute('val', $xmlnsMain, 'Table of Contents')
        [ref] $null = $docPartObj.AppendChild($XmlDocument.CreateElement('w', 'docPartUnique', $xmlnsMain))
        [ref] $null = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdEndPr', $xmlnsMain))

        $sdtContent = $sdt.AppendChild($XmlDocument.CreateElement('w', 'stdContent', $xmlnsMain))
        $p1 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
        $pPr1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
        $pStyle1 = $pPr1.AppendChild($XmlDocument.CreateElement('w', 'pStyle', $xmlnsMain))
        [ref] $null = $pStyle1.SetAttribute('val', $xmlnsMain, 'TOC')
        $r1 = $p1.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        $t1 = $r1.AppendChild($XmlDocument.CreateElement('w', 't', $xmlnsMain))
        [ref] $null = $t1.AppendChild($XmlDocument.CreateTextNode($TOC.Name))

        $p2 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
        $pPr2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
        $tabs2 = $pPr2.AppendChild($XmlDocument.CreateElement('w', 'tabs', $xmlnsMain))
        $tab2 = $tabs2.AppendChild($XmlDocument.CreateElement('w', 'tab', $xmlnsMain))
        [ref] $null = $tab2.SetAttribute('val', $xmlnsMain, 'right')
        [ref] $null = $tab2.SetAttribute('leader', $xmlnsMain, 'dot')
        [ref] $null = $tab2.SetAttribute('pos', $xmlnsMain, '9016')
        #10790?!
        $r2 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        ##TODO: Refactor duplicate code
        $fldChar1 = $r2.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
        [ref] $null = $fldChar1.SetAttribute('fldCharType', $xmlnsMain, 'begin')

        $r3 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        $instrText = $r3.AppendChild($XmlDocument.CreateElement('w', 'instrText', $xmlnsMain))
        [ref] $null = $instrText.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        [ref] $null = $instrText.AppendChild($XmlDocument.CreateTextNode(' TOC \o "1-3" \h \z \u '))

        $r4 = $p2.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        $fldChar2 = $r4.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
        [ref] $null = $fldChar2.SetAttribute('fldCharType', $xmlnsMain, 'separate')

        $p3 = $sdtContent.AppendChild($XmlDocument.CreateElement('w', 'p', $xmlnsMain))
        $r5 = $p3.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        #$rPr3 = $r3.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain));
        $fldChar3 = $r5.AppendChild($XmlDocument.CreateElement('w', 'fldChar', $xmlnsMain))
        [ref] $null = $fldChar3.SetAttribute('fldCharType', $xmlnsMain, 'end')

        return $sdt
    }
}

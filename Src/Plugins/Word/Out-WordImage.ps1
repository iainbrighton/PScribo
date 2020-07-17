function Out-WordImage
{
<#
    .SYNOPSIS
        Output Image to Word.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Image,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlnswpDrawingWordProcessing = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        $xmlnsDrawingMain = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        $xmlnsDrawingPicture = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        $xmlnsRelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

        $p = $XmlDocument.CreateElement('w', 'p', $xmlnsMain)
        $pPr = $p.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)

        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))
        [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Image.Align.ToLower())

        $r = $p.AppendChild($XmlDocument.CreateElement('w', 'r', $xmlnsMain))
        [ref] $null = $r.AppendChild($XmlDocument.CreateElement('w', 'rPr', $xmlnsMain))
        $drawing = $r.AppendChild($XmlDocument.CreateElement('w', 'drawing', $xmlnsMain))
        $inline = $drawing.AppendChild($XmlDocument.CreateElement('wp', 'inline', $xmlnswpDrawingWordProcessing))
        [ref] $null = $inline.SetAttribute('distT', '0')
        [ref] $null = $inline.SetAttribute('distB', '0')
        [ref] $null = $inline.SetAttribute('distL', '0')
        [ref] $null = $inline.SetAttribute('distR', '0')

        $extent = $inline.AppendChild($XmlDocument.CreateElement('wp', 'extent', $xmlnswpDrawingWordProcessing))
        [ref] $null = $extent.SetAttribute('cx', $Image.WidthEm)
        [ref] $null = $extent.SetAttribute('cy', $Image.HeightEm)

        $effectExtent = $inline.AppendChild($XmlDocument.CreateElement('wp', 'effectExtent', $xmlnswpDrawingWordProcessing))
        [ref] $null = $effectExtent.SetAttribute('l', '0')
        [ref] $null = $effectExtent.SetAttribute('t', '0')
        [ref] $null = $effectExtent.SetAttribute('r', '0')
        [ref] $null = $effectExtent.SetAttribute('b', '0')

        $docPr = $inline.AppendChild($XmlDocument.CreateElement('wp', 'docPr', $xmlnswpDrawingWordProcessing))
        [ref] $null = $docPr.SetAttribute('id', $Image.ImageNumber)
        [ref] $null = $docPr.SetAttribute('name', $Image.Name)
        [ref] $null = $docPr.SetAttribute('descr', $Image.Name)

        $cNvGraphicFramePr = $inline.AppendChild($XmlDocument.CreateElement('wp', 'cNvGraphicFramePr', $xmlnswpDrawingWordProcessing))
        $graphicFrameLocks = $cNvGraphicFramePr.AppendChild($XmlDocument.CreateElement('a', 'graphicFrameLocks', $xmlnsDrawingMain))
        [ref] $null = $graphicFrameLocks.SetAttribute('noChangeAspect', '1')

        $graphic = $inline.AppendChild($XmlDocument.CreateElement('a', 'graphic', $xmlnsDrawingMain))
        $graphicData = $graphic.AppendChild($XmlDocument.CreateElement('a', 'graphicData', $xmlnsDrawingMain))
        [ref] $null = $graphicData.SetAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

        $pic = $graphicData.AppendChild($XmlDocument.CreateElement('pic', 'pic', $xmlnsDrawingPicture))
        $nvPicPr = $pic.AppendChild($XmlDocument.CreateElement('pic', 'nvPicPr', $xmlnsDrawingPicture))
        $cNvPr = $nvPicPr.AppendChild($XmlDocument.CreateElement('pic', 'cNvPr', $xmlnsDrawingPicture))
        [ref] $null = $cNvPr.SetAttribute('id', $Image.ImageNumber)
        [ref] $null = $cNvPr.SetAttribute('name', $Image.Name)
        [ref] $null = $cNvPr.SetAttribute('descr', $Image.Name)
        $cNvPicPr = $nvPicPr.AppendChild($XmlDocument.CreateElement('pic', 'cNvPicPr', $xmlnsDrawingPicture))
        $picLocks = $cNvPicPr.AppendChild($XmlDocument.CreateElement('a', 'picLocks', $xmlnsDrawingMain))
        [ref] $null = $picLocks.SetAttribute('noChangeAspect', '1')
        [ref] $null = $picLocks.SetAttribute('noChangeArrowheads', '1')

        $blipFill = $pic.AppendChild($XmlDocument.CreateElement('pic', 'blipFill', $xmlnsDrawingPicture))
        $blip = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'blip', $xmlnsDrawingMain))
        [ref] $null = $blip.SetAttribute('embed', $xmlnsRelationships, $Image.Name)
        [ref] $null = $blip.SetAttribute('cstate', 'print')
        $extlst = $blip.AppendChild($XmlDocument.CreateElement('a', 'extlst', $xmlnsDrawingMain))
        $ext = $extlst.AppendChild($XmlDocument.CreateElement('a', 'ext', $xmlnsDrawingMain))
        [ref] $null = $ext.SetAttribute('uri', '')
        [ref] $null = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'srcRect', $xmlnsDrawingMain))
        $stretch = $blipFill.AppendChild($XmlDocument.CreateElement('a', 'stretch', $xmlnsDrawingMain))
        [ref] $null = $stretch.AppendChild($XmlDocument.CreateElement('a', 'fillRect', $xmlnsDrawingMain))

        $spPr = $pic.AppendChild($XmlDocument.CreateElement('pic', 'spPr', $xmlnsDrawingPicture))
        [ref] $null = $spPr.SetAttribute('bwMode', 'auto')
        $xfrm = $spPr.AppendChild($XmlDocument.CreateElement('a', 'xfrm', $xmlnsDrawingMain))
        $off = $xfrm.AppendChild($XmlDocument.CreateElement('a', 'off', $xmlnsDrawingMain))
        [ref] $null = $off.SetAttribute('x', '0')
        [ref] $null = $off.SetAttribute('y', '0')
        $ext = $xfrm.AppendChild($XmlDocument.CreateElement('a', 'ext', $xmlnsDrawingMain))
        [ref] $null = $ext.SetAttribute('cx', $Image.WidthEm)
        [ref] $null = $ext.SetAttribute('cy', $Image.HeightEm)

        $prstGeom = $spPr.AppendChild($XmlDocument.CreateElement('a', 'prstGeom', $xmlnsDrawingMain))
        [ref] $null = $prstGeom.SetAttribute('prst', 'rect')
        [ref] $null = $prstGeom.AppendChild($XmlDocument.CreateElement('a', 'avLst', $xmlnsDrawingMain))

        [ref] $null = $spPr.AppendChild($XmlDocument.CreateElement('a', 'noFill', $xmlnsDrawingMain))

        $ln = $spPr.AppendChild($XmlDocument.CreateElement('a', 'ln', $xmlnsDrawingMain))
        [ref] $null = $ln.AppendChild($XmlDocument.CreateElement('a', 'noFill', $xmlnsDrawingMain))

        return $p
    }
}

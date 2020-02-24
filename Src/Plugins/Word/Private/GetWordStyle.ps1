function GetWordStyle
{
<#
    .SYNOPSIS
        Generates Word Xml style element from a PScribo document style.
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        ## PScribo document style
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Style,

        [Parameter(Mandatory)]
        [System.Xml.XmlDocument] $XmlDocument,

        [Parameter(Mandatory)]
        [ValidateSet('Paragraph', 'Character')]
        [System.String] $Type
    )
    process
    {
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        if ($Type -eq 'Paragraph')
        {
            $styleId = $Style.Id
            $styleName = $Style.Name
            $linkId = '{0}Char' -f $Style.Id
        }
        else
        {
            $styleId = '{0}Char' -f $Style.Id
            $styleName = '{0} Char' -f $Style.Name
            $linkId = $Style.Id
        }

        $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlnsMain)
        [ref] $null = $documentStyle.SetAttribute('type', $xmlnsMain, $Type.ToLower())

        if ($Style.Id -eq $Document.DefaultStyle)
        {
            ## Set as default style
            [ref] $null = $documentStyle.SetAttribute('default', $xmlnsMain, 1)
            $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'uiPriority', $xmlnsMain))
            [ref] $null = $uiPriority.SetAttribute('val', $xmlnsMain, 1)
        }
        elseif ($Style.Hidden -eq $true)
        {
            ## Semi hide style (headers and footers etc)
            [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain))
        }
        elseif (($Document.TableStyles.Values | ForEach-Object -Process {

                    $_.HeaderStyle
                    $_.RowStyle
                    $_.AlternateRowStyle
                }) -contains $Style.Id) {
            ## Semi hide styles behind table styles (except default style!)
            [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlnsMain))
        }

        [ref] $null = $documentStyle.SetAttribute('styleId', $xmlnsMain, $styleId)
        $documentStyleName = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlnsMain))
        [ref] $null = $documentStyleName.SetAttribute('val', $xmlnsMain, $styleName)
        $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlnsMain))
        [ref] $null = $basedOn.SetAttribute('val', $xmlnsMain, 'Normal')
        $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'link', $xmlnsMain))
        [ref] $null = $link.SetAttribute('val', $xmlnsMain, $linkId)
        $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'next', $xmlnsMain))
        [ref] $null = $next.SetAttribute('val', $xmlnsMain, 'Normal')
        [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlnsMain))
        $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlnsMain))
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepNext', $xmlnsMain))
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlnsMain))
        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlnsMain))
        [ref] $null = $spacing.SetAttribute('before', $xmlnsMain, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlnsMain, 0)
        ## Set the <w:jc> (justification) element
        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlnsMain))

        if ($Style.Align.ToLower() -eq 'justify')
        {
            [ref] $null = $jc.SetAttribute('val', $xmlnsMain, 'distribute')
        }
        else
        {
            [ref] $null = $jc.SetAttribute('val', $xmlnsMain, $Style.Align.ToLower())
        }

        if ($Style.BackgroundColor)
        {
            $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlnsMain))
            [ref] $null = $shd.SetAttribute('val', $xmlnsMain, 'clear')
            [ref] $null = $shd.SetAttribute('color', $xmlnsMain, 'auto')
            [ref] $null = $shd.SetAttribute('fill', $xmlnsMain, (ConvertToWordColor -Color $Style.BackgroundColor))
        }
        [ref] $null = $documentStyle.AppendChild((GetWordStyleRunPr -Style $Style -XmlDocument $XmlDocument))

        return $documentStyle
    }
}

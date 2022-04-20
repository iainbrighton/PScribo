function Get-WordStyle
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
    begin
    {
        ## Semi hide styles behind table styles (except default style!)
        $hiddenStyleIds = $Document.TableStyles.Values |
            ForEach-Object -Process { $_.HeaderStyle; $_.RowStyle; $_.AlternateRowStyle }
    }
    process
    {
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

        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $documentStyle = $XmlDocument.CreateElement('w', 'style', $xmlns)
        [ref] $null = $documentStyle.SetAttribute('type', $xmlns, $Type.ToLower())
        [ref] $null = $documentStyle.SetAttribute('styleId', $xmlns, $styleId)

        if ($Style.Id -eq $Document.DefaultStyle)
        {
            ## Set as default style
            [ref] $null = $documentStyle.SetAttribute('default', $xmlns, 1)
            $uiPriority = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'uiPriority', $xmlns))
            [ref] $null = $uiPriority.SetAttribute('val', $xmlns, 1)
        }
        elseif ($Style.Hidden -eq $true)
        {
            ## Semi hide style (headers and footers etc)
            [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlns))
        }
        elseif ($hiddenStyleIds -contains $Style.Id)
        {
            [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'semiHidden', $xmlns))
        }

        $documentStyleName = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'name', $xmlns))
        [ref] $null = $documentStyleName.SetAttribute('val', $xmlns, $styleName)

        $basedOn = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'basedOn', $xmlns))
        [ref] $null = $basedOn.SetAttribute('val', $xmlns, 'Normal')

        $link = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'link', $xmlns))
        [ref] $null = $link.SetAttribute('val', $xmlns, $linkId)

        $next = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'next', $xmlns))
        [ref] $null = $next.SetAttribute('val', $xmlns, 'Normal')

        [ref] $null = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'qFormat', $xmlns))

        $pPr = $documentStyle.AppendChild($XmlDocument.CreateElement('w', 'pPr', $xmlns))
        [ref] $null = $pPr.AppendChild($XmlDocument.CreateElement('w', 'keepLines', $xmlns))

        $spacing = $pPr.AppendChild($XmlDocument.CreateElement('w', 'spacing', $xmlns))
        [ref] $null = $spacing.SetAttribute('before', $xmlns, 0)
        [ref] $null = $spacing.SetAttribute('after', $xmlns, 0)

        $jc = $pPr.AppendChild($XmlDocument.CreateElement('w', 'jc', $xmlns))
        if ($Style.Align.ToLower() -eq 'justify')
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, 'distribute')
        }
        else
        {
            [ref] $null = $jc.SetAttribute('val', $xmlns, $Style.Align.ToLower())
        }

        if ($Style.BackgroundColor)
        {
            $backgroundColor = ConvertTo-WordColor -Color (Resolve-PScriboStyleColor -Color $Style.BackgroundColor)
            $shd = $pPr.AppendChild($XmlDocument.CreateElement('w', 'shd', $xmlns))
            [ref] $null = $shd.SetAttribute('val', $xmlns, 'clear')
            [ref] $null = $shd.SetAttribute('color', $xmlns, 'auto')
            [ref] $null = $shd.SetAttribute('fill', $xmlns, $backgroundColor)
        }

        $rPr = Get-WordStyleRunPr -Style $Style -XmlDocument $XmlDocument
        [ref] $null = $documentStyle.AppendChild($rPr)

        return $documentStyle
    }
}

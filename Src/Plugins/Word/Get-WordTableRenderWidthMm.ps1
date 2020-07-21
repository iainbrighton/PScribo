function Get-WordTableRenderWidthMm
{
<#
    .SYNOPSIS
        Calcualtes a table's (maximum) rendering size in millimetres.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 100)]
        [System.UInt16] $TableWidth,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateRange(0, 10)]
        [System.UInt16] $Tabs,

        [Parameter(Mandatory, ValueFromPipeline)]
        [System.String] $Orientation
    )
    process
    {
        if ($TableWidth -eq 0)
        {
            ## Word will autofit contents as necessary, but LibreOffice won't
            ## so we just have assume that we're using all available width
            $TableWidth = 100
        }
        if ($Orientation -eq 'Portrait')
        {
            $pageWidthMm = $Document.Options['PageWidth'] - ($Document.Options['MarginLeft'] + $Document.Options['MarginRight'])
        }
        elseif ($Orientation -eq 'Landscape')
        {
            $pageWidthMm = $Document.Options['PageHeight'] - ($Document.Options['MarginLeft'] + $Document.Options['MarginRight'])
        }
        $indentWidthMm = ConvertTo-Mm -Point ($Tabs * 36)
        $tableWidthRenderMm = (($pageWidthMm / 100) * $TableWidth) + $indentWidthMm
        if ($tableWidthRenderMm -gt $pageWidthMm)
        {
            ## We now need to deal with tables being pushed outside the page margin so just return the maximum permitted
            $tableWidthRenderMm = $pageWidthMm - $indentWidthMm
            Write-PScriboMessage -Message ($localized.TableWidthOverflowWarning -f $tableWidthRenderMm) -IsWarning
        }
        return $tableWidthRenderMm
    }
}

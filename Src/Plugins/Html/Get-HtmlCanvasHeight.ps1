function Get-HtmlCanvasHeight
{
<#
    .SYNOPSIS
        Calculates the usable document canvas height.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [System.String] $Orientation
    )
    process
    {
        if ($Orientation -eq 'Portrait')
        {
            $pageHeight = $Document.Options['PageHeight']
            return ($pageHeight - $Document.Options['MarginTop'] - $Document.Options['MarginBottom']) -as [System.Int32]
        }
        else
        {
            $pageHeight = $Document.Options['PageWidth']
            return ($pageHeight - $Document.Options['MarginTop'] - (($Document.Options['MarginBottom'] * 1.5) -as [System.Int32]))
        }
    }
}

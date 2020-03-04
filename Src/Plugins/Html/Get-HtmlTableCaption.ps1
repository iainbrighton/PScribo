function Get-HtmlTableCaption
{
<#
    .SYNOPSIS
        Generates html <p> caption from a PScribo.Table object.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style
        $pscriboParagraph = [PSCustomObject] @{
            Id                = '{0}{1}' -f $tableStyle.CaptionPrefix, $Table.Number
            Text              = '{0} {1} {2}' -f $tableStyle.CaptionPrefix, $Table.Number, $Table.Caption
            Type              = 'PScribo.Paragraph'
            Style             = $tableStyle.CaptionStyle
            Value             = $null
            NewLine           = $false
            Tabs              = 0
            Bold              = $false
            Italic            = $false
            Underline         = $false
            Font              = $null
            Size              = $null
            Color             = $null
            Orientation       = $script:currentOrientation
            IsSectionBreakEnd = $false
        }
        return (Out-HtmlParagraph -Paragraph $pscriboParagraph)
    }
}

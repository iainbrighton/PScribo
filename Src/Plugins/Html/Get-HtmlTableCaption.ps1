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

        ## Scaffold paragraph and paragraph run for table caption
        $paragraphId = '{0}{1}' -f $tableStyle.CaptionPrefix, $Table.Number
        $paragraph = New-PScriboParagraph -Id $paragraphId -Style $tableStyle.CaptionStyle -NoIncrementCounter
        $paragraphRunText = '{0} {1} {2}' -f $tableStyle.CaptionPrefix, $Table.CaptionNumber, $Table.Caption
        $paragraphRun = New-PScriboParagraphRun -Text $paragraphRunText
        $paragraphRun.IsParagraphRunEnd = $true
        [ref] $null = $paragraph.Sections.Add($paragraphRun)

        return (Out-HtmlParagraph -Paragraph $paragraph)
    }
}

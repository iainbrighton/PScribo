function Get-MarkdownTableCaption
{
<#
    .SYNOPSIS
        Generates Markdown table caption from a PScribo.Table object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','Options')]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboMarkdownOption
        }
    }
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

        return (Out-MarkdownParagraph -Paragraph $paragraph)
    }
}

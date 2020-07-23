function Out-MarkdownParagraph
{
<#
    .SYNOPSIS
        Output formatted paragraph run.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Paragraph
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
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        $paragraphInlineStyle = Get-MarkdownParagraphStyle -Paragraph $Paragraph
        [ref] $null = $paragraphBuilder.Append($paragraphInlineStyle)

        foreach ($paragraphRun in $Paragraph.Sections)
        {
            $inlineStyle = Get-MarkdownParagraphRunInlineStyle -ParagraphRun $paragraphRun
            [ref] $null = $paragraphBuilder.Append($inlineStyle)

            ## Create a separate text block for each line/break
            $text = Resolve-PScriboToken -InputObject $paragraphRun.Text
            [ref] $null = $paragraphBuilder.Append($text)

            if (($paragraphRun.IsParagraphRunEnd -eq $false) -and
                ($paragraphRun.NoSpace -eq $false))
                {
                [ref] $null = $paragraphBuilder.Append(' ')
            }

            [ref] $null = $paragraphBuilder.Append($inlineStyle)
        }

        [ref] $null = $paragraphBuilder.Append($paragraphInlineStyle)
        $script:currentPScriboObject = 'PScribo.Paragraph'

        $convertToAlignedStringParams = @{
            InputObject = $paragraphBuilder.ToString()
            Width       = $options.TextWidth
            Align       = 'Left'
            Markdown    = $true
        }
        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

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
        $convertToAlignedStringParams = @{
            Width    = $options.TextWidth
            Align    = 'Left'
            Markdown = $true
        }

        $getMarkdownParagraphInlineStyleParams = @{ }
        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $getMarkdownParagraphInlineStyleParams['Style'] = $Paragraph.Style
        }

        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        foreach ($paragraphRun in $Paragraph.Sections)
        {
            $getMarkdownParagraphInlineStyleParams['ParagraphRun'] = $paragraphRun
            $inlineStyle = Get-MarkdownParagraphRunInlineStyle @getMarkdownParagraphInlineStyleParams
            [ref] $null = $paragraphBuilder.Append($inlineStyle)

            ## Create a separate text block for each line/break
            $text = Resolve-PScriboToken -InputObject $paragraphRun.Text
            [ref] $null = $paragraphBuilder.Append($text)

            # $lines = $text -split '\r\n?|\n'
            # for ($l = 0; $l -lt $lines.Count; $l++)
            # {
            #     $convertToAlignedStringParams['InputObject'] = $lines[$l]
            #     $line = ConvertTo-AlignedString @convertToAlignedStringParams
            #     [ref] $null = $paragraphBuilder.Append($line)
#
            #     ## Don't add a line break (double space) to the last line/break
            #     if ($l -lt ($lines.Count - 1))
            #     {
            #         [ref] $null = $paragraphBuilder.AppendLine('  ')
            #     }
            # }

            if (($paragraphRun.IsParagraphRunEnd -eq $false) -and
                ($paragraphRun.NoSpace -eq $false))
            {
                [ref] $null = $paragraphBuilder.Append(' ')
            }

            [ref] $null = $paragraphBuilder.Append($inlineStyle)
        }

        #return $paragraphBuilder.ToString()
        $convertToAlignedStringParams['InputObject'] = $paragraphBuilder.ToString()
        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

function Get-MarkdownDocument
{
<#
    .SYNOPSIS
        Markdown output plugin for PScribo.

    .DESCRIPTION
        Renders a GFM markdown file representation of a PScribo document object.

    .NOTES
        Enables unit testing without writing .md file to disk.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','Options')]
    param
    (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $Options = New-PScriboMarkdownOption
        }
    }
    process
    {
        [System.Text.StringBuilder] $markdownBuilder = New-Object -Type 'System.Text.StringBuilder'

        foreach ($subSection in $Document.Sections.GetEnumerator())
        {
            $currentIndentationLevel = 1
            if ($null -ne $subSection.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $subSection.Level +1
            }
            Write-PScriboProcessSectionId -SectionId $subSection.Id -SectionType $subSection.Type -Indent $currentIndentationLevel

            switch ($subSection.Type)
            {
                'PScribo.Section'
                {
                    $markdownSection = Out-MarkdownSection -Section $subSection
                    [ref] $null = $markdownBuilder.Append($markdownSection)
                }
                'PScribo.Paragraph'
                {
                    $markdownParagraph = Out-MarkdownParagraph -Paragraph $subSection
                    [ref] $null = $markdownBuilder.Append($markdownParagraph)
                }
                'PScribo.PageBreak'
                {
                    $markdownPageBreak = Out-MarkdownPageBreak
                    [ref] $null = $markdownBuilder.Append($markdownPageBreak)
                }
                'PScribo.LineBreak'
                {
                    $markdownLineBreak = Out-MarkdownLineBreak
                    [ref] $null = $markdownBuilder.Append($markdownLineBreak)
                }
                'PScribo.Table'
                {
                    $markdownTables = Out-MarkdownTable -Table $subSection
                    [ref] $null = $markdownBuilder.Append($markdownTables)
                }
                'PScribo.TOC'
                {
                    $markdownTOC = Out-MarkdownTOC -TOC $subSection
                    [ref] $null = $markdownBuilder.Append($markdownTOC)
                }
                'PScribo.BlankLine'
                {
                    $blankline = Out-MarkdownBlankLine -BlankLine $subSection
                    [ref] $null = $markdownBuilder.Append($blankline)
                }
                'PScribo.Image'
                {
                    $markdownImage = Out-MarkdownImage -Image $subSection
                    [ref] $null = $markdownBuilder.Append($markdownImage)
                }
                Default
                {
                    Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        ## Write binary image data to the bottom of the document
        foreach ($image in (Get-PSCriboImage -Section $Document.Sections))
        {
            if (($null -ne $Options) -and ($Options['EmbedImage'] -eq $true)) # -or
            {
                if ($script:currentPScriboObject -eq 'PScribo.Paragraph')
                {
                    [ref] $null = $markdownBuilder.AppendLine()
                    $script:currentPScriboObject = 'PScribo.Document'
                }
                $imageBase64 = [System.Convert]::ToBase64String($image.Bytes)
                [ref] $null = $markdownBuilder.AppendFormat('[ref_{0}]: data:{1};base64,{2}', $image.Name.ToLower(), $image.MIMEType, $imageBase64).AppendLine()
            }
        }

        return $markdownBuilder.ToString()
    }
}

function Out-MarkdownDocument
{
<#
    .SYNOPSIS
        Markdown output plugin for PScribo.

    .DESCRIPTION
        Outputs a CommonMark markdown file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    param
    (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    process
    {
        $pluginName = 'Markdown'
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Document.Name)

        ## Merge the document, text default and specified text options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboMarkdownOption
            DocumentOptions = $Document.Options
            PluginOptions = $Options
        }
        $Options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams
        $script:currentPageNumber = 1

        [System.Text.StringBuilder] $textBuilder = New-Object -Type 'System.Text.StringBuilder'

        # $firstPageHeader = Out-TextHeaderFooter -Header -FirstPage
        # [ref] $null = $textBuilder.Append($firstPageHeader)

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
                    [ref] $null = $textBuilder.Append($markdownSection)
                }
                'PScribo.Paragraph'
                {
                    $markdownParagraph = Out-MarkdownParagraph -Paragraph $subSection
                    [ref] $null = $textBuilder.Append($markdownParagraph)
                }
                'PScribo.PageBreak'
                {
                    $markdownPageBreak = Out-MarkdownPageBreak
                    [ref] $null = $textBuilder.Append($markdownPageBreak)
                }
                'PScribo.LineBreak'
                {
                    $markdownLineBreak = Out-MarkdownLineBreak
                    [ref] $null = $textBuilder.Append($markdownLineBreak)
                }
                'PScribo.Table'
                {
                    $markdownTables = Out-MarkdownTable -Table $subSection
                    [ref] $null = $textBuilder.Append($markdownTables)
                }
                'PScribo.TOC'
                {
                    $markdownTOC = Out-MarkdownTOC -TOC $subSection
                    [ref] $null = $textBuilder.Append($markdownTOC)
                }
                'PScribo.BlankLine'
                {
                    $blankline = Out-MarkdownBlankLine -BlankLine $subSection
                    [ref] $null = $textBuilder.Append($blankline)
                }
                'PScribo.Image'
                {
                    $markdownImage = Out-MarkdownImage -Image $subSection
                    [ref] $null = $textBuilder.Append($markdownImage)
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
            $imageBase64 = [System.Convert]::ToBase64String($image.Bytes)
            [ref] $null = $textBuilder.AppendFormat('[image_ref_{0}]: data:{1};base64,{2}', $image.Name.ToLower(), $image.MIMEType, $imageBase64).AppendLine()
            # [image_ref_a32ff4ads]: data:image/png;base64,iVBORw0KGgoAAAANSUhEke02C1MyA29UWKgPA...RS12D==
        }

        # $pageFooter =Out-TextHeaderFooter -Footer
        # [ref] $null = $textBuilder.Append($pageFooter)

        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)
        $destinationPath = Join-Path -Path $Path ('{0}.md' -f $Document.Name)
        Write-PScriboMessage -Message ($localized.SavingFile -f $destinationPath)
        Set-Content -Value ($textBuilder.ToString()) -Path $destinationPath -Encoding $Options.Encoding
        [ref] $null = $textBuilder

        if ($stopwatch.Elapsed.TotalSeconds -gt 90)
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeMinutes -f $stopwatch.Elapsed.TotalMinutes)
        }
        else
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeSeconds -f $stopwatch.Elapsed.TotalSeconds)
        }

        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath)
    }
}

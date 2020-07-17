function Out-TextDocument
{
<#
    .SYNOPSIS
        Text output plugin for PScribo.

    .DESCRIPTION
        Outputs a text file representation of a PScribo document object.
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
        $pluginName = 'Text'
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Document.Name)

        ## Merge the document, text default and specified text options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboTextOption
            DocumentOptions = $Document.Options
            PluginOptions = $Options
        }
        $Options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams
        $script:currentPageNumber = 1

        [System.Text.StringBuilder] $textBuilder = New-Object -Type 'System.Text.StringBuilder'
        $firstPageHeader = Out-TextHeaderFooter -Header -FirstPage
        [ref] $null = $textBuilder.Append($firstPageHeader)

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
                    [ref] $null = $textBuilder.Append((Out-TextSection -Section $subSection))
                }
                'PScribo.Paragraph'
                {
                    [ref] $null = $textBuilder.Append((Out-TextParagraph -Paragraph $subSection))
                }
                'PScribo.PageBreak'
                {
                    [ref] $null = $textBuilder.Append((Out-TextPageBreak))
                }
                'PScribo.LineBreak'
                {
                    [ref] $null = $textBuilder.Append((Out-TextLineBreak))
                }
                'PScribo.Table'
                {
                    [ref] $null = $textBuilder.Append((Out-TextTable -Table $subSection))
                }
                'PScribo.TOC'
                {
                    [ref] $null = $textBuilder.Append((Out-TextTOC -TOC $subSection))
                }
                'PScribo.BlankLine'
                {
                    [ref] $null = $textBuilder.Append((Out-TextBlankLine -BlankLine $subSection))
                }
                'PScribo.Image'
                {
                    [ref] $null = $textBuilder.Append((Out-TextImage -Image $subSection))
                }
                Default
                {
                    Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        $pageFooter =Out-TextHeaderFooter -Footer
        [ref] $null = $textBuilder.Append($pageFooter)

        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)
        $destinationPath = Join-Path -Path $Path ('{0}.txt' -f $Document.Name)
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

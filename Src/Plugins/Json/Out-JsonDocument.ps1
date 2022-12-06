function Out-JsonDocument
{
<#
    .SYNOPSIS
        Json output plugin for PScribo.

    .DESCRIPTION
        Outputs a Json file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    param
    (
        ## ThePScribo document object to convert to a Json document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ## Output directory path for the .json file
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
        $pluginName = 'Json'
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Document.Name)

        ## Merge the document, plugin default and specified/specific plugin options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboJsonOption
            DocumentOptions = $Document.Options
            PluginOptions = $Options
        }
        $Options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams
        $script:currentPageNumber = 1

        ## Initializing JSON object
        $jsonBuilder = [ordered]@{}

        ## Initializing paragraph counter
        [int]$paragraph = 1

        ## Initializing table counter
        [int]$table = 1

        ## Generating header
        $header = Out-JsonHeaderFooter -Header -FirstPage
        if ($null -ne $header) {
            [ref] $null = $jsonBuilder.Add("header", $header)
            [ref] $null = $header
        }

        foreach ($subSection in $Document.Sections.GetEnumerator())
        {
            # Write-Host "Type: $($subSection.Type)"
            switch ($subSection.Type)
            {
                'PScribo.Section'
                {
                    ## Corrects behavior where NOTOC* heading is used
                    if (("" -eq $subSection.Number))
                    {
                        [ref] $null = $jsonBuilder.Add($subSection.Name, (Out-JsonSection -Section $subSection))
                    }
                    else
                    {
                        [ref] $null = $jsonBuilder.Add($subSection.Number, (Out-JsonSection -Section $subSection))
                    }
                }
                'PScribo.Paragraph'
                {
                    [ref] $null = $jsonBuilder.Add("paragraph$($paragraph)", (Out-JsonParagraph -Paragraph $subSection))
                    $paragraph++
                }
                'PScribo.Table'
                {
                    [ref] $null = $jsonBuilder.Add("table$($table)", (Out-JsonTable -Table $subSection))
                    $table++
                }
                'PScribo.TOC'
                {
                    [ref] $null = $jsonBuilder.Add("toc", (Out-JsonTOC -TOC $subSection))
                }
                Default
                {
                    Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        ## Generating footer
        $footer = Out-JsonHeaderFooter -Footer
        if ($null -ne $footer) {
            [ref] $null = $jsonBuilder.Add("footer", $footer)
            [ref] $null = $footer
        }

        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)
        $destinationPath = Join-Path -Path $Path ('{0}.json' -f $Document.Name)
        Write-PScriboMessage -Message ($localized.SavingFile -f $destinationPath)
        $jsonBuilder | ConvertTo-Json -Depth 100 | Set-Content -Path $destinationPath -Encoding $Options.Encoding
        [ref] $null = $jsonBuilder

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

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
        $script:currentPScriboObject = 'PScribo.Document'
        $markdownDocument = Get-MarkdownDocument -Document $Document -Options $Options

        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)

        $destinationPath = Join-Path -Path $Path ('{0}.md' -f $Document.Name)
        Write-PScriboMessage -Message ($localized.SavingFile -f $destinationPath)
        Set-Content -Value $markdownDocument -Path $destinationPath -Encoding $Options.Encoding

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

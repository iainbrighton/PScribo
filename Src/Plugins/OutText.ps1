function OutText {
<#
    .SYNOPSIS
        Text output plugin for PScribo.
    .DESCRIPTION
        Outputs a text file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    param (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,

        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        $pluginName = 'Text';

        <#! OutText.Internal.ps1 !#>

    }
    process {

        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);

        ## Merge the document, text default and specified text options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboTextOption;
            DocumentOptions = $Document.Options;
            PluginOptions = $Options;
        }
        $Options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

        [System.Text.StringBuilder] $textBuilder = New-Object System.Text.StringBuilder;
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $textBuilder.Append((OutTextSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $textBuilder.Append(($s | OutTextParagraph)); }
                'PScribo.PageBreak' { [ref] $null = $textBuilder.AppendLine((OutTextPageBreak)); }
                'PScribo.LineBreak' { [ref] $null = $textBuilder.AppendLine((OutTextLineBreak)); }
                'PScribo.Table' { [ref] $null = $textBuilder.AppendLine(($s | OutTextTable)); }
                'PScribo.TOC' { [ref] $null = $textBuilder.AppendLine(($s | OutTextTOC)); }
                'PScribo.BlankLine' { [ref] $null = $textBuilder.AppendLine(($s | OutTextBlankLine)); }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path -Path $Path ('{0}.txt' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        Set-Content -Value ($textBuilder.ToString()) -Path $destinationPath -Encoding $Options.Encoding;
        [ref] $null = $textBuilder;
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);

    } #end process
} #end function OutText

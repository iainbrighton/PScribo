function OutHtml {
<#
    .SYNOPSIS
        Html output plugin for PScribo.
    .DESCRIPTION
        Outputs a Html file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.IO.FileInfo])]
    param (
        ## PScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject] $Document,

        ## Output directory path for the .html file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        $pluginName = 'Html';

        <#! OutHtml.Internal.ps1 !#>

    } #end begin
    process {

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);

        ## Merge the document, plugin default and specified/specific plugin options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboHtmlOption;
            DocumentOptions = $Document.Options;
            PluginOptions = $Options;
        }
        $options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

        $noPageLayoutStyle = $Options['NoPageLayoutStyle'];
        $topMargin = ConvertMmToEm -Millimeter $options['MarginTop'];
        $leftMargin = ConvertMmToEm -Millimeter $options['MarginLeft'];
        $bottomMargin = ConvertMmToEm -Millimeter $options['MarginBottom'];
        $rightMargin = ConvertMmToEm -Millimeter $options['MarginRight'];

        [System.Text.StringBuilder] $htmlBuilder = New-Object System.Text.StringBuilder;
        [ref] $null = $htmlBuilder.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">');
        [ref] $null = $htmlBuilder.AppendLine('<html xmlns="http://www.w3.org/1999/xhtml">');
        [ref] $null = $htmlBuilder.AppendLine('<head><title>{0}</title>' -f $Document.Name);
        [ref] $null = $htmlBuilder.AppendLine('{0}</head><body><page>' -f (OutHtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles -NoPageLayoutStyle:$noPageLayoutStyle));
        [ref] $null = $htmlBuilder.AppendFormat('<div class="{0}" style="padding-top: {1}em; padding-left: {2}em; padding-bottom: {3}em; padding-right: {4}em;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine();
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $htmlBuilder.Append((OutHtmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $htmlBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                'PScribo.Table' { [ref] $null = $htmlBuilder.Append((OutHtmlTable -Table $s)); }
                'PScribo.LineBreak' { [ref] $null = $htmlBuilder.Append((OutHtmlLineBreak)); }
                'PScribo.PageBreak' { [ref] $null = $htmlBuilder.Append((OutHtmlPageBreak)); } ## Page breaks are implemented as line breaks with extra padding
                'PScribo.TOC' { [ref] $null = $htmlBuilder.Append((OutHtmlTOC -TOC $s)); }
                'PScribo.BlankLine' { [ref] $null = $htmlBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            } #end switch
        } #end foreach section
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path $Path ('{0}.html' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        $htmlBuilder.ToString().TrimEnd() | Out-File -FilePath $destinationPath -Force -Encoding utf8;
        [ref] $null = $htmlBuilder;
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        Write-Output (Get-Item -Path $destinationPath);

    } #end process
} #end function OutHtml

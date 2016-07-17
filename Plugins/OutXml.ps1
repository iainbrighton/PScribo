function OutXml {
<#
    .SYNOPSIS
        Xml output plugin for PScribo.
    .DESCRIPTION
        Outputs a xml representation of a PScribo document object.
#>
    [CmdletBinding()]
    param (
        ## ThePScribo document object to convert to a xml document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,

        ## Output directory path for the .xml file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        <#! OutXml.Internal.ps1 !#>

    }
    process {

        $pluginName = 'Xml';
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew();
        WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name);
        $documentName = $Document.Name;

        $xmlDocument = New-Object -TypeName System.Xml.XmlDocument;
        [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
        $documentId = ($Document.Id -replace '[^a-z0-9-_\.]','').ToLower();
        $element = $xmlDocument.AppendChild($xmlDocument.CreateElement($documentId));
        [ref] $null = $element.SetAttribute("name", $documentName);
        foreach ($s in $Document.Sections.GetEnumerator()) {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}[..]' -f $s.Id.Substring(0,36); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type) {
                'PScribo.Section' { [ref] $null = $element.AppendChild((OutXmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s)); }
                'PScribo.Table' { [ref] $null = $element.AppendChild((OutXmlTable -Table $s)); }
                'PScribo.PageBreak'{ } ## Page breaks are not implemented for Xml output
                'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                'PScribo.TOC' { } ## TOC is not implemented for Xml output
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning;
                }
            } #end switch
        } #end foreach
        $stopwatch.Stop();
        WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name);
        $destinationPath = Join-Path $Path ('{0}.xml' -f $Document.Name);
        WriteLog -Message ($localized.SavingFile -f $destinationPath);
        $xmlDocument.Save($destinationPath);
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        ## Return the file reference to the pipeline
        Write-Output (Get-Item -Path $destinationPath);

    } #end process
} #end function outxml

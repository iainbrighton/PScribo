function OutXmlSection
{
<#
    .SYNOPSIS
        Output formatted Xml section.
#>
    [CmdletBinding()]
    param
    (
        ## PScribo document section
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Section
    )
    process
    {
        $sectionId = ($Section.Id -replace '[^a-z0-9-_\.]','').ToLower();
        $element = $xmlDocument.CreateElement('section');
        [ref] $null = $element.SetAttribute("name", $Section.Name);
        foreach ($s in $Section.Sections.GetEnumerator())
        {
            if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
            else { $sectionId = $s.Id; }
            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type)
            {
                'PScribo.Section' { [ref] $null = $element.AppendChild((OutXmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s)); }
                'PScribo.Table' { [ref] $null = $element.AppendChild((OutXmlTable -Table $s)); }
                'PScribo.PageBreak' { } ## Page breaks are not implemented for Xml output
                'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                'PScribo.TOC' { } ## TOC is not implemented for Xml output
                'PScribo.Image' { [ref] $null = $element.AppendChild((OutXmlImage -Image $s)); }
                Default {
                    WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning;
                }
            }
        }
        return $element;
    }
}

function OutHtmlSection
{
<#
    .SYNOPSIS
        Output formatted Html section.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        ## Section to output
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Section
    )
    process
    {
        [System.Text.StringBuilder] $sectionBuilder = New-Object System.Text.StringBuilder;
        if ($Section.IsSectionBreak)
        {
            [ref] $null = $sectionBuilder.Append((OutHtmlPageBreak -Orientation $Section.Orientation));
        }
        $encodedSectionName = [System.Net.WebUtility]::HtmlEncode($Section.Name);
        if ($Document.Options['EnableSectionNumbering']) { [System.String] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectionName; }
        else { [System.String] $sectionName = '{0}' -f $encodedSectionName; }
        [int] $headerLevel = $Section.Number.Split('.').Count;

        ## Html <h5> is the maximum supported level
        if ($headerLevel -gt 6)
        {
            WriteLog -Message $localized.MaxHeadingLevelWarning -IsWarning;
            $headerLevel = 6;
        }

        if ([System.String]::IsNullOrEmpty($Section.Style))
        {
            $className = $Document.DefaultStyle;
        }
        else
        {
            $className = $Section.Style;
        }

        if ($Section.Tabs -gt 0)
        {
            $tabEm = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Section.Tabs)) -Format 'f2';
            [ref] $null = $sectionBuilder.AppendFormat('<div style="margin-left: {0}rem;">' -f $tabEm);
        }
        [ref] $null = $sectionBuilder.AppendFormat('<a name="{0}"><h{1} class="{2}">{3}</h{1}></a>', $Section.Id, $headerLevel, $className, $sectionName.TrimStart());
        if ($Section.Tabs -gt 0)
        {
            [ref] $null = $sectionBuilder.Append('</div>');
        }

        foreach ($s in $Section.Sections.GetEnumerator())
        {
            if ($s.Id.Length -gt 40)
            {
                $sectionId = '{0}[..]' -f $s.Id.Substring(0,36);
            }
            else
            {
                $sectionId = $s.Id;
            }

            $currentIndentationLevel = 1;
            if ($null -ne $s.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $s.Level +1;
            }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
            switch ($s.Type)
            {
                'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutHtmlSection -Section $s)); }
                'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append((OutHtmlParagraph -Paragraph $s)); }
                'PScribo.LineBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlLineBreak)); }
                'PScribo.PageBreak' { [ref] $null = $sectionBuilder.Append((OutHtmlPageBreak -Orientation $Section.Orientation)); }
                'PScribo.Table' { [ref] $null = $sectionBuilder.Append((OutHtmlTable -Table $s)); }
                'PScribo.BlankLine' { [ref] $null = $sectionBuilder.Append((OutHtmlBlankLine -BlankLine $s)); }
                'PScribo.Image' { [ref] $null = $sectionBuilder.Append((OutHtmlImage -Image $s)); }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning; }
            }
        }
        return $sectionBuilder.ToString();
    }
}

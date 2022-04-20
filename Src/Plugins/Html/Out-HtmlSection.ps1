function Out-HtmlSection
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
        [System.Text.StringBuilder] $sectionBuilder = New-Object System.Text.StringBuilder
        if ($Section.IsSectionBreak)
        {
            [ref] $null = $sectionBuilder.Append((Out-HtmlPageBreak -Orientation $Section.Orientation))
        }
        $encodedSectionName = [System.Net.WebUtility]::HtmlEncode($Section.Name)
        if ($Document.Options['EnableSectionNumbering'])
        {
            [System.String] $sectionName = '{0} {1}' -f $Section.Number, $encodedSectionName
        }
        else
        {
            [System.String] $sectionName = '{0}' -f $encodedSectionName
        }
        [int] $headerLevel = $Section.Number.Split('.').Count

        ## Html <h5> is the maximum supported level
        if ($headerLevel -gt 6)
        {
            Write-PScriboMessage -Message $localized.MaxHeadingLevelWarning -IsWarning
            $headerLevel = 6
        }

        if ([System.String]::IsNullOrEmpty($Section.Style))
        {
            $className = $Document.DefaultStyle
        }
        else
        {
            $className = $Section.Style
        }

        if ($Section.Tabs -gt 0)
        {
            $tabEm = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter (12.7 * $Section.Tabs)) -Format 'f2'
            [ref] $null = $sectionBuilder.AppendFormat('<div style="margin-left: {0}rem;">' -f $tabEm)
        }
        [ref] $null = $sectionBuilder.AppendFormat('<a name="{0}"><h{1} class="{2}">{3}</h{1}></a>', $Section.Id, $headerLevel, $className, $sectionName.TrimStart())
        if ($Section.Tabs -gt 0)
        {
            [ref] $null = $sectionBuilder.Append('</div>')
        }

        foreach ($subSection in $Section.Sections.GetEnumerator())
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
                    [ref] $null = $sectionBuilder.Append((Out-HtmlSection -Section $subSection))
                }
                'PScribo.Paragraph' {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlParagraph -Paragraph $subSection))
                }
                'PScribo.LineBreak'
                {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlLineBreak))
                }
                'PScribo.PageBreak'
                {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlPageBreak -Orientation $Section.Orientation))
                }
                'PScribo.Table'
                {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlTable -Table $subSection))
                }
                'PScribo.BlankLine'
                {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlBlankLine -BlankLine $subSection))
                }
                'PScribo.Image'
                {
                    [ref] $null = $sectionBuilder.Append((Out-HtmlImage -Image $subSection))
                }
                Default
                {
                     Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        return $sectionBuilder.ToString()
    }
}

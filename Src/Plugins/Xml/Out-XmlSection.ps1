function Out-XmlSection
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
        $sectionId = ($Section.Id -replace '[^a-z0-9-_\.]','').ToLower()
        $element = $xmlDocument.CreateElement('section')
        [ref] $null = $element.SetAttribute("name", $Section.Name)

        foreach ($subSection in $Section.Sections.GetEnumerator())
        {
            $currentIndentationLevel = 1
            if ($null -ne $subSection.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $subSection.Level +1
            }
            Write-PScriboProcessSectionId -SectionId $sectionId -SectionType $subSection.Type -Indent $currentIndentationLevel

            switch ($subSection.Type)
            {
                'PScribo.Section'
                {
                    [ref] $null = $element.AppendChild((Out-XmlSection -Section $subSection))
                }
                'PScribo.Paragraph'
                {
                    [ref] $null = $element.AppendChild((Out-XmlParagraph -Paragraph $subSection))
                }
                'PScribo.Table'
                {
                    [ref] $null = $element.AppendChild((Out-XmlTable -Table $subSection))
                }
                'PScribo.PageBreak' { } ## Page breaks are not implemented for Xml output
                'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                'PScribo.TOC' { } ## TOC is not implemented for Xml output
                'PScribo.Image'
                {
                    [ref] $null = $element.AppendChild((Out-XmlImage -Image $subSection))
                }
                Default
                {
                    Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        }

        return $element
    }
}

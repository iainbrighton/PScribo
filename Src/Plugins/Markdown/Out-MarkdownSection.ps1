function Out-MarkdownSection
{
<#
    .SYNOPSIS
        Output formatted markdown section.
#>
    [CmdletBinding()]
    param
    (
        ## Section to output
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Section
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboMarkdownOption
        }
    }
    process
    {
        $sectionLevel = $Section.Level +1
        if ($sectionLevel -gt 6)
        {
            $sectionLevel = 6
        }
        $sectionLeader = ''.PadRight($sectionLevel, '#')
        $sectionBuilder = New-Object -TypeName System.Text.StringBuilder
        if ($Document.Options['EnableSectionNumbering'])
        {
            [string] $sectionName = '{0} {1} {2}' -f $sectionLeader, $Section.Number, $Section.Name
        }
        else
        {
            [string] $sectionName = '{0} {1}' -f $sectionLeader, $Section.Name
        }
        [ref] $null = $sectionBuilder.AppendLine($sectionName).AppendLine()

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
                    $markdownSection = Out-MarkdownSection -Section $subSection
                    [ref] $null = $sectionBuilder.Append($markdownSection)
                }
                'PScribo.Paragraph'
                {
                    $markdownParagraph = Out-MarkdownParagraph -Paragraph $subSection
                    [ref] $null = $sectionBuilder.Append($markdownParagraph)
                }
                'PScribo.PageBreak'
                {
                    $markdownPageBreak = Out-MarkdownPageBreak
                    [ref] $null = $sectionBuilder.Append($markdownPageBreak)
                }
                'PScribo.LineBreak'
                {
                    $markdownLineBreak = Out-MarkdownLineBreak
                    [ref] $null = $sectionBuilder.Append($markdownLineBreak)
                }
                'PScribo.Table'
                {
                    $markdownTables = Out-MarkdownTable -Table $subSection
                    [ref] $null = $sectionBuilder.Append($markdownTables)
                }
                'PScribo.BlankLine'
                {
                    $blankline = Out-MarkdownBlankLine -BlankLine $subSection
                    [ref] $null = $sectionBuilder.Append($blankline)
                }
                'PScribo.Image'
                {
                    $markdownImage = Out-MarkdownImage -Image $subSection
                    [ref] $null = $sectionBuilder.Append($markdownImage)
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

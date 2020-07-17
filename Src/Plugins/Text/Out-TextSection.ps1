function Out-TextSection
{
<#
    .SYNOPSIS
        Output formatted text section.
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
            $options = New-PScriboTextOption
        }
    }
    process
    {
        $padding = ''.PadRight(($Section.Tabs * 4), ' ')
        $sectionBuilder = New-Object -TypeName System.Text.StringBuilder
        if ($Document.Options['EnableSectionNumbering'])
        {
            [string] $sectionName = '{0} {1}' -f $Section.Number, $Section.Name
        }
        else
        {
            [string] $sectionName = '{0}' -f $Section.Name
        }
        [ref] $null = $sectionBuilder.AppendLine()
        [ref] $null = $sectionBuilder.Append($padding)
        [ref] $null = $sectionBuilder.AppendLine($sectionName.TrimStart())
        [ref] $null = $sectionBuilder.Append($padding)
        [ref] $null = $sectionBuilder.AppendLine(''.PadRight(($options.SeparatorWidth - $padding.Length), $options.SectionSeparator))

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
                    [ref] $null = $sectionBuilder.Append((Out-TextSection -Section $subSection))
                }
                'PScribo.Paragraph'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextParagraph -Paragraph $subSection))
                }
                'PScribo.PageBreak'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextPageBreak)) ## Page breaks implemented as line break with extra padding
                }
                'PScribo.LineBreak'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextLineBreak))
                }
                'PScribo.Table'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextTable -Table $subSection))
                }
                'PScribo.BlankLine'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextBlankLine -BlankLine $subSection))
                }
                'PScribo.Image'
                {
                    [ref] $null = $sectionBuilder.Append((Out-TextImage -Image $subSection))
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

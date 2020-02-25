function OutTextSection
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
        foreach ($s in $Section.Sections.GetEnumerator())
        {
            if ($s.Id.Length -gt 40)
            {
                $sectionId = '{0}..' -f $s.Id.Substring(0,38)
            }
            else
            {
                $sectionId = $s.Id
            }
            $currentIndentationLevel = 1
            if ($null -ne $s.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $s.Level +1
            }
            WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel
            switch ($s.Type)
            {
                'PScribo.Section' { [ref] $null = $sectionBuilder.Append((OutTextSection -Section $s)) }
                'PScribo.Paragraph' { [ref] $null = $sectionBuilder.Append(($s | OutTextParagraph)) }
                'PScribo.PageBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextPageBreak)) }  ## Page breaks implemented as line break with extra padding
                'PScribo.LineBreak' { [ref] $null = $sectionBuilder.AppendLine((OutTextLineBreak)) }
                'PScribo.Table' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextTable)) }
                'PScribo.BlankLine' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextBlankLine)) }
                'PScribo.Image' { [ref] $null = $sectionBuilder.AppendLine(($s | OutTextImage)) }
                Default { WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning }
            }
        }
        return $sectionBuilder.ToString()
    }
}

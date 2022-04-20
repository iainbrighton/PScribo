function Out-TextTOC
{
<#
    .SYNOPSIS
        Output formatted Table of Contents
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TOC
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
        $tocBuilder = New-Object -TypeName System.Text.StringBuilder
        [ref] $null = $tocBuilder.AppendLine($TOC.Name)
        [ref] $null = $tocBuilder.AppendLine(''.PadRight($options.SeparatorWidth, $options.SectionSeparator))

        if ($Options.ContainsKey('EnableSectionNumbering'))
        {
            $maxSectionNumberLength = $Document.TOC.Number | ForEach-Object { $_.Length } | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
            foreach ($tocEntry in $Document.TOC)
            {
                $sectionNumberPaddingLength = $maxSectionNumberLength - $tocEntry.Number.Length
                $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ')
                $sectionPadding = ''.PadRight($sectionNumberPaddingLength, ' ')
                [ref] $null = $tocBuilder.AppendFormat('{0}{1}  {2}{3}', $tocEntry.Number, $sectionPadding, $sectionNumberIndent, $tocEntry.Name).AppendLine()
            }
        }
        else
        {
            $maxSectionNumberLength = $Document.TOC.Level | Sort-Object | Select-Object -Last 1
            foreach ($tocEntry in $Document.TOC)
            {
                $sectionNumberIndent = ''.PadRight($tocEntry.Level, ' ')
                [ref] $null = $tocBuilder.AppendFormat('{0}{1}', $sectionNumberIndent, $tocEntry.Name).AppendLine()
            }
        }

        return $tocBuilder.ToString()
    }
}

function Get-WordNumberStyle
{
<#
    .SYNOPSIS
        Outputs Office Open XML numbering document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlElement])]
    param
    (
        ## PScribo document styles
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Xml.XmlDocument] $XmlDocument
    )
    begin
    {
        $bulletMatrix = @{
            Disc   = [PSCustomObject] @{ Text = ''; Font = 'Symbol'; }
            Circle = [PSCustomObject] @{ Text = 'o'; Font = 'Courier New'; }
            Square = [PSCustomObject] @{ Text = ''; Font = 'Wingdings'; }
            Dash   = [PSCustomObject] @{ Text = '-'; Font = 'Courier New' }
        }
    }
    process
    {
        ## Create the Numbering.xml document
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $xmlnsmc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

        $abstractNum = $xmlDocument.CreateElement('w', 'abstractNum', $xmlns)
        [ref] $null = $abstractNum.SetAttribute('abstractNumId', $xmlns, $list.Number -1)
        $multiLevelType = $abstractNum.AppendChild($xmlDocument.CreateElement('w', 'multiLevelType', $xmlns))
        [ref] $null = $multiLevelType.SetAttribute('val', $xmlns, 'hybridMultilevel')

        $listLevels = Get-WordListLevel -List $List
        foreach ($level in ($listLevels.Keys | Sort-Object))
        {
            $listLevel = $listLevels[$level]
            if ($listLevel.IsNumbered)
            {
                $numberStyle = $Document.NumberStyles[$listLevel.NumberStyle]
            }

            $lvl = $abstractNum.AppendChild($xmlDocument.CreateElement('w', 'lvl', $xmlns))
            [ref] $null = $lvl.SetAttribute('ilvl', $xmlns, $level)
            $start = $lvl.AppendChild($xmlDocument.CreateElement('w', 'start', $xmlns))
            [ref] $null = $start.SetAttribute('val', $xmlns, 1)

            if ($listLevel.IsNumbered)
            {
                switch ($numberStyle.Format)
                {
                    Number
                    {
                        $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                        $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                        [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'decimal')
                        [ref] $null = $lvlText.SetAttribute('val', $xmlns, ('%{0}{1}' -f ($level +1), $numberStyle.Suffix))
                    }
                    Letter
                    {
                        $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                        $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                        if ($numberStyle.Uppercase)
                        {
                            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'upperLetter')
                        }
                        else
                        {
                            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'lowerLetter')
                        }
                        [ref] $null = $lvlText.SetAttribute('val', $xmlns, ('%{0}{1}' -f ($level +1), $numberStyle.Suffix))
                    }
                    Roman
                    {
                        $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                        $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                        if ($numberStyle.Uppercase)
                        {
                            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'upperRoman')
                        }
                        else
                        {
                            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'lowerRoman')
                        }
                        [ref] $null = $lvlText.SetAttribute('val', $xmlns, ('%{0}{1}' -f ($level +1), $numberStyle.Suffix))
                    }
                    Custom
                    {
                        $regexMatch = [Regex]::Match($numberStyle.Custom, '%+')
                        if ($regexMatch.Success -eq $true)
                        {
                            $numberString = $regexMatch.Value
                            if ($numberString.Length -eq 1)
                            {
                                $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                                $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                                [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'decimal')
                            }
                            elseif ($numberString.Length -eq 2)
                            {
                                $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                                $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                                [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'decimalZero')
                            }
                            else
                            {
                                $AlternateContent = $lvl.AppendChild($xmlDocument.CreateElement('mc', 'AlternateContent', $xmlnsmc))
                                $Choice = $AlternateContent.AppendChild($xmlDocument.CreateElement('mc', 'Choice', $xmlnsmc))
                                [ref] $null = $Choice.SetAttribute('Requires', 'w14')
                                $numFmtC = $Choice.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                                [ref] $null = $numFmtC.SetAttribute('val', $xmlns, 'custom')

                                if ($numberString.Length -eq 3)
                                {
                                    [ref] $null = $numFmtC.SetAttribute('format', $xmlns, '001, 002, 003, ...')
                                }
                                elseif ($numberString.Length -eq 4)
                                {
                                    [ref] $null = $numFmtC.SetAttribute('format', $xmlns, '0001, 0002, 0003, ...')
                                }
                                elseif ($numberString.Length -ge 5)
                                {
                                    [ref] $null = $numFmtC.SetAttribute('format', $xmlns, '00001, 00002, 00003, ...')
                                }

                                $Fallback = $AlternateContent.AppendChild($xmlDocument.CreateElement('mc', 'Fallback', $xmlnsmc))
                                $numFmtF = $Fallback.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                                [ref] $null = $numFmtF.SetAttribute('val', $xmlns, 'decimal')

                                $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                            }
                        }
                        else
                        {
                            Write-PScriboMessage 'Invalid custom number format' -IsWarning
                            $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                            [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'none')
                        }
                        $replacementString = '%{0}' -f ($level +1)
                        $wordNumberString = $numberStyle.Custom -replace '%+', $replacementString
                        [ref] $null = $lvlText.SetAttribute('val', $xmlns, $wordNumberString)
                    }
                }
            }
            else
            {
                $numFmt = $lvl.AppendChild($xmlDocument.CreateElement('w', 'numFmt', $xmlns))
                [ref] $null = $numFmt.SetAttribute('val', $xmlns, 'bullet')
                $lvlText = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlText', $xmlns))
                [ref] $null = $lvlText.SetAttribute('val', $xmlns, $bulletMatrix[$listLevel.BulletStyle].Text)
                $rPr = $lvl.AppendChild($xmlDocument.CreateElement('w', 'rPr', $xmlns))
                $rFonts = $rPr.AppendChild($xmlDocument.CreateElement('w', 'rFonts', $xmlns))
                [ref] $null = $rFonts.SetAttribute('ascii', $xmlns, $bulletMatrix[$listLevel.BulletStyle].Font)
                [ref] $null = $rFonts.SetAttribute('hAnsi', $xmlns, $bulletMatrix[$listLevel.BulletStyle].Font)
                [ref] $null = $rFonts.SetAttribute('hint', $xmlns, 'default')
            }

            $lvlJc = $lvl.AppendChild($xmlDocument.CreateElement('w', 'lvlJc', $xmlns))

            if ($listLevel.IsNumbered)
            {
                $indent = $numberStyle.Indent
                $hanging = $numberStyle.Hanging

                if ($numberStyle.Align -eq 'Left')
                {
                    [ref] $null = $lvlJc.SetAttribute('val', $xmlns,'start')
                    if ($indent -eq 0)
                    {
                        $indent = ($level + 1) * 680
                    }
                    if ($hanging -eq 0)
                    {
                        $hanging = 400
                    }
                }
                elseif ($numberStyle.Align -eq 'Right')
                {
                    [ref] $null = $lvlJc.SetAttribute('val', $xmlns, 'end')
                    if ($indent -eq 0)
                    {
                        $indent = ($level + 1) * 680
                    }
                    if ($hanging -eq 0)
                    {
                        $hanging = 120
                    }
                }
            }
            else
            {
                [ref] $null = $lvlJc.SetAttribute('val', $xmlns, 'Right')
                $indent = ($level + 1) * 640
                $hanging = 280
            }

            $pPr = $lvl.AppendChild($xmlDocument.CreateElement('w', 'pPr', $xmlns))
            $ind = $pPr.AppendChild($xmlDocument.CreateElement('w', 'ind', $xmlns))

            [ref] $null = $ind.SetAttribute('left', $xmlns, $indent)
            [ref] $null = $ind.SetAttribute('hanging', $xmlns, $hanging)
        }

        return $abstractNum
    }
}

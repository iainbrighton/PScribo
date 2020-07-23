$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    function GetMatch
    {
        [CmdletBinding()]
        param
        (
            [System.String] $String
        )
        Write-Verbose "Pre Match : '$String'"
        $matchString = $String.Replace('/','\/')
        if (-not $String.StartsWith('^'))
        {
            $matchString = $matchString.Replace('[..]','[\s\S]+')
            $matchString = $matchString.Replace('[??]','([\s\S]+)?')
        }
        Write-Verbose "Post Match: '$matchString'"
        return $matchString
    }

    Describe 'Plugins\Markdown\Out-MarkdownTable' {

        $Document = Document -Name 'TestDocument' -ScriptBlock { }
        $pscriboDocument = $Document
        $script:currentPScriboObject = 'PScribo.Document'
        $script:currentPageNumber = 1

        Context 'Table' {

            It 'outputs table headers' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet

                $table = $processes | Table
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('^\| Id \| Name \| WorkingSet {0}' -f [System.Environment]::NewLine)
                $result | Should Match ('\| :- \| :--- \| :--------- {0}' -f [System.Environment]::NewLine)
            }

            It 'outputs table rows' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet

                $table = $processes | Table
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('(\| [\s\w]+ \| [\s\w]+ \| [\s\w]+ {0}){{{1}}}' -f [System.Environment]::NewLine, $processes.Count)
            }

            It 'ends with blank/empty line' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet

                $table = $processes | Table
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('{0}{0}$' -f [System.Environment]::NewLine)
            }

            It 'outputs table caption after (by default)' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet
                $testCaption = '- Test Caption'

                $table = $processes | Table -Caption $testCaption
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('{0}{0}_Table {1} {2}_  {0}{0}$' -f [System.Environment]::NewLine, $table.CaptionNumber, $testCaption)
            }

            It 'outputs table caption before table when specified' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet
                $tableAboveStyleParams = @{
                    Id                = 'TableAbove'
                    BorderWidth       = 1
                    BorderColor       = '2a70be'
                    HeaderStyle       = 'TableDefaultHeading'
                    RowStyle          = 'TableDefaultRow'
                    AlternateRowStyle = 'TableDefaultAltRow'
                    CaptionStyle      = 'Caption'
                    CaptionLocation   = 'Above'
                }
                $testCaption = '- Test Caption'

                TableStyle @tableAboveStyleParams -Default -Verbose:$false
                $table = $processes | Table -Style 'TableAbove' -Caption $testCaption
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('^_Table {0} {1}_  {2}{2}\|' -f $table.CaptionNumber, $testCaption, [System.Environment]::NewLine)
            }

        }

        Context 'List' {

            It 'outputs table (header) per row' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet

                $table = $processes | Table -List
                $result = Out-MarkdownTable -Table $table

                $expected = GetMatch ('\| :-+ \| :-+[..]\| :-+ \| :-+[..]\| :-+ \| :-+')
                $result | Should Match $expected
            }

            It 'ends with blank/empty line' {

                $processes = Get-Process | Select-Object -First 3 -Property Id,Name,WorkingSet

                $table = $processes | Table -List
                $result = Out-MarkdownTable -Table $table

                $result | Should Match ('{0}{0}$' -f [System.Environment]::NewLine)
            }
        }
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginsRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginsRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;



InModuleScope 'PScribo' {

    function GetMatch
    {
        [CmdletBinding()]
        param
        (
            [System.String] $String,
            [System.Management.Automation.SwitchParameter] $Complete
        )
        Write-Verbose "Pre Match : '$String'"
        $matchString = $String.Replace('/','\/')
        if (-not $String.StartsWith('^'))
        {
            $matchString = $matchString.Replace('[..]','[\s\S]+')
            $matchString = $matchString.Replace('[??]','([\s\S]+)?')
            if ($Complete)
            {
                $matchString = '^<w:test xmlns:w="http:\/\/schemas.openxmlformats.org\/wordprocessingml\/2006\/main">{0}<\/w:test>$' -f $matchString
            }
        }
        Write-Verbose "Post Match: '$matchString'"
        return $matchString
    }

    Describe 'Plugins\Word\Out-WordBlankLine' {

        It 'appends paragraph "<w:p />"' {
            $document = Document -Name 'TestDocument' {
                BlankLine
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p />'
            $testDocument.DocumentElement.OuterXml | Should Match $expected
        }

        It 'appends paragraph "<w:p />" per blankline' {
            $document = Document -Name 'TestDocument' {
                BlankLine -Count 2
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p /><w:p />'
            $testDocument.DocumentElement.OuterXml | Should Match $expected
        }
    }
}

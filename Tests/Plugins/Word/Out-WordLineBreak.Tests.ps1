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

    Describe 'Plugins\Word\Out-WordLineBreak' {

        It 'outputs paragraph properties "<w:p><w:pPr><w:pBdr>[..]</w:pBdr></w:pPr></w:p>"' {
            $document = Document -Name 'TestDocument' {
                LineBreak
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:p><w:pPr><w:pBdr>[..]</w:pBdr></w:pPr></w:p>'
            $testDocument.OuterXml  | Should Match $expected
        }

        It 'outputs border "<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto" /></w:pBdr>"' {
            $document = Document -Name 'TestDocument' {
                LineBreak
            }

            $testDocument = Get-WordDocument -Document $document

            $expected = GetMatch '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto" /></w:pBdr>'
            $testDocument.OuterXml  | Should Match $expected
        }
    }
}

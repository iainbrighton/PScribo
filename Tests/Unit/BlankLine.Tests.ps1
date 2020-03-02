$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'BlankLine' {
        $pscriboDocument = Document 'ScaffoldDocument' {}

        It 'returns a PSCustomObject object' {
            $b = BlankLine

            $b.GetType().Name | Should Be 'PSCustomObject'
        }

        It 'creates a PScribo.BlankLine type' {
            $b = BlankLine

            $b.Type | Should Be 'PScribo.BlankLine'
        }

        It 'creates blank line with no parameters' {
            $b = BlankLine

            $b.LineCount | Should Be 1
        }

        It 'creates blank line by named -Count parameter' {
            $count = 2

            $b = BlankLine -Count $count

            $b.LineCount | Should Be $count
        }
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'PageBreak' {

        $pscriboDocument = Document 'ScaffoldDocument' {}

        It 'returns a PSCustomObject object' {
            $p = PageBreak

            $p.GetType().Name | Should Be 'PSCustomObject'
        }

        It 'creates a PScribo.PageBreak type' {
            $p = PageBreak

            $p.Type | Should Be 'PScribo.PageBreak'
        }

        It 'creates page break with no parameters' {
            $p = PageBreak

            $p.Id | Should Not Be $null
        }

        It 'creates page break by named -Id parameter' {
            $id = 'Test'

            $p = PageBreak -Id $id

            $p.Id | Should Be $id
        }

        It 'creates page break by positional -Id parameter' {
            $id = 'Test'

            $p = PageBreak $id

            $p.Id | Should Be $id
        }
    }
}

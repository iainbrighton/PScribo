$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'LineBreak' {
        $pscriboDocument = Document 'ScaffoldDocument' {};

        It 'returns a PSCustomObject object' {
            $l = LineBreak

            $l.GetType().Name | Should Be 'PSCustomObject'
        }

        It 'creates a PScribo.LineBreak type' {
            $l = LineBreak

            $l.Type | Should Be 'PScribo.LineBreak'
        }

        It 'creates line break with no parameters' {
            $l = LineBreak

            $l.Id | Should Not Be $null
        }

        It 'creates line break by named -Id parameter' {
            $id = 'Test'

            $l = LineBreak -Id $id

            $l.Id | Should Be $id
        }

        It 'creates line break by positional -Id parameter' {
            $id = 'Test'

            $l = LineBreak $id

            $l.Id | Should Be $id
        }
    }
}

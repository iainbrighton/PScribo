$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginsRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginsRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Word\ConvertTo-WordColor' {

        It 'converts to "abcdef" to "ABCDEF"' {
            $result = ConvertTo-WordColor 'abcdef'

            $result | Should BeExactly 'ABCDEF'
        }

        It 'converts "#abcdef" to "ABCDEF"' {
            $result = ConvertTo-WordColor '#abcdef'

            $result | Should BeExactly 'ABCDEF'
        }

        It 'converts "abc" to "AABBCC"' {
            $result = ConvertTo-WordColor 'abc'

            $result | Should BeExactly 'AABBCC'
        }

        It 'converts "#abc" to "AABBCC"' {
            $result = ConvertTo-WordColor '#abc'

            $result | Should BeExactly 'AABBCC'
        }
    }
}

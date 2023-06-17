$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    Describe -Name 'NumberStyle' {

        BeforeEach {

            $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock { }
        }

        It 'returns "PSCustomObject" object' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number

            $pscriboDocument.NumberStyles[$numberStyleName].GetType().Name | Should -eq 'PSCustomObject'
        }

        It 'defaults to "." suffix' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number

            $pscriboDocument.NumberStyles[$numberStyleName].Suffix | Should -eq '.'
        }

        It 'defaults to "Right" alignment' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number

            $pscriboDocument.NumberStyles[$numberStyleName].Align | Should -eq 'Right'
        }

        It 'defaults to "Lowercase"' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number

            $pscriboDocument.NumberStyles[$numberStyleName].Uppercase | Should -eq $false
        }

        It 'creates "Letter" number style' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Letter

            $pscriboDocument.NumberStyles[$numberStyleName].Format | Should -eq 'Letter'
        }

        It 'creates "Roman" number style' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Roman

            $pscriboDocument.NumberStyles[$numberStyleName].Format | Should -eq 'Roman'
        }

        It 'creates "Custom" number style' {
            $numberStyleName = 'Custom'
            $customFormat = 'ABC%'

            NumberStyle -Name $numberStyleName -Custom $customFormat

            $pscriboDocument.NumberStyles[$numberStyleName].Format | Should -eq 'Custom'
            $pscriboDocument.NumberStyles[$numberStyleName].Custom | Should -eq $customFormat
        }

        It 'sets custom suffix' {
            $numberStyleName = 'Custom'
            $customSuffix = ':'

            NumberStyle -Name $numberStyleName -Format Number -Suffix $customSuffix

            $pscriboDocument.NumberStyles[$numberStyleName].Suffix | Should -eq $customSuffix
        }

        It 'sets left alignment' {
            $numberStyleName = 'Custom'
            $customAlignment = 'Left'

            NumberStyle -Name $numberStyleName -Format Number -Align $customAlignment

            $pscriboDocument.NumberStyles[$numberStyleName].Align | Should -eq $customAlignment
        }

        It 'sets uppercase text' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number -Uppercase

            $pscriboDocument.NumberStyles[$numberStyleName].Uppercase | Should -eq $true
        }

        It 'sets Word indent level when specified' {
            $numberStyleName = 'Custom'
            $customIndent = 500

            NumberStyle -Name $numberStyleName -Format Number -Indent $customIndent

            $pscriboDocument.NumberStyles[$numberStyleName].Indent | Should -eq $customIndent
        }

        It 'sets Word hanging level when specified' {
            $numberStyleName = 'Custom'
            $customHanging = 250

            NumberStyle -Name $numberStyleName -Format Number -Hanging $customHanging

            $pscriboDocument.NumberStyles[$numberStyleName].Hanging | Should -eq $customHanging
        }

        It 'sets default document number style' {
            $numberStyleName = 'Custom'

            NumberStyle -Name $numberStyleName -Format Number -Default

            $pscriboDocument.DefaultNumberStyle | Should Be $numberStyleName
        }

    }
}

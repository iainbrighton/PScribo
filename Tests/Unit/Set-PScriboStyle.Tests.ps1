$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {


    Describe 'Set-PScriboStyle' {

        $pscriboDocument = Document 'ScaffoldDocument' {}
        Style -Name 'MyCustomStyle' -Size 16

        Context 'By Row' {

            It 'sets row style by reference' {
                ($service = Get-Process | Select-Object -Last 1) |
                    Set-Style -Style 'MyCustomStyle'

                $service.__Style | Should Be 'MyCustomStyle'
            }
            It 'sets row style by pipeline' {
                $service = Get-Process | Select-Object -Last 1 |
                    Set-Style -Style 'MyCustomStyle' -PassThru

                $service.__Style | Should Be 'MyCustomStyle'
            }
        }

        Context 'By Cell' {

            It 'sets cell style on a single property by reference' {
                ($service = Get-Process | Select-Object -Last 1) |
                    Set-Style -Style 'MyCustomStyle' -Property SI

                $service.SI__Style | Should Be 'MyCustomStyle'
            }

            It 'sets cell style on a single property by pipeline' {
                $service = Get-Process | Select-Object -Last 1 |
                    Set-Style -Style 'MyCustomStyle' -Property SI -PassThru

                $service.SI__Style | Should Be 'MyCustomStyle'
            }
            It 'sets cell style on a multiple properties by reference' {
                ($service = Get-Process | Select-Object -Last 1) |
                    Set-Style -Style 'MyCustomStyle' -Property SI,ProcessName

                $service.SI__Style | Should Be 'MyCustomStyle'
                $service.ProcessName__Style | Should Be 'MyCustomStyle'
            }

            It 'sets cell style on a multiple properties by pipeline' {
                $service = Get-Process | Select-Object -Last 1 |
                    Set-Style -Style 'MyCustomStyle' -Property SI,ProcessName -PassThru

                $service.SI__Style | Should Be 'MyCustomStyle'
                $service.ProcessName__Style | Should Be 'MyCustomStyle'
            }
        }
    }
}

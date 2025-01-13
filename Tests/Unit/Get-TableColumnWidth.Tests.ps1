$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Get-TableColumnWidth' {

        Context 'By named parameter' {
            BeforeAll {
                $testData = @(
                    [PSCustomObject]@{
                        Name = 'Test1'
                        Description = 'Description 1'
                        Value = 123
                    }
                )
            }

            It 'accepts single object input' {
                $result = Get-TableColumnWidth -InputObject $testData[0]
                $result.Count | Should -BeExactly 3
            }

            It 'accepts array input and uses first item' {
                $result = Get-TableColumnWidth -InputObject $testData
                $result.Count | Should -BeExactly 3
            }

            It 'throws when custom widths exceed 100' {
                $customWidths = @{
                    'Name' = 60
                    'Description' = 50
                }
                { Get-TableColumnWidth -InputObject $testData -CustomWidths $customWidths } | 
                Should -Throw "*exceeds the total available width*"
            }

            It 'distributes width evenly with no custom widths' {
                $result = Get-TableColumnWidth -InputObject $testData
                $sum = ($result | Measure-Object -Sum).Sum
                $sum | Should -BeExactly 100
                $result[0] | Should -BeExactly 33
                $result[1] | Should -BeExactly 33
                $result[2] | Should -BeExactly 34
            }

            It 'respects custom widths and distributes remaining width' {
                $customWidths = @{
                    'Name' = 50
                }
                $result = Get-TableColumnWidth -InputObject $testData -CustomWidths $customWidths
                $sum = ($result | Measure-Object -Sum).Sum
                $sum | Should -BeExactly 100
                $result[0] | Should -BeExactly 50
                $result[1] | Should -BeExactly 25
                $result[2] | Should -BeExactly 25
            }

            It 'handles custom widths for all columns' {
                $customWidths = @{
                    'Name' = 30
                    'Description' = 40
                    'Value' = 30
                }
                $result = Get-TableColumnWidth -InputObject $testData -CustomWidths $customWidths
                $sum = ($result | Measure-Object -Sum).Sum
                $sum | Should -BeExactly 100
                $result[0] | Should -BeExactly 30
                $result[1] | Should -BeExactly 40
                $result[2] | Should -BeExactly 30
            }

            It 'handles single column input' {
                $singleColumn = [PSCustomObject]@{ Name = 'Test' }
                $result = Get-TableColumnWidth -InputObject $singleColumn
                $result.Count | Should -BeExactly 1
                $result[0] | Should -BeExactly 100
            }

            It 'distributes remaining width when custom widths total less than 100' {
                $customWidths = @{
                    'Name' = 20
                    'Description' = 20
                    'Value' = 20
                }
                $result = Get-TableColumnWidth -InputObject $testData -CustomWidths $customWidths
                $sum = ($result | Measure-Object -Sum).Sum
                $sum | Should -BeExactly 100
                $result[0] | Should -BeExactly 33
                $result[1] | Should -BeExactly 33
                $result[2] | Should -BeExactly 34
            }
        }

        Context 'Input validation' {
            It 'throws on null input' {
                { Get-TableColumnWidth -InputObject $null } | Should -Throw
            }

            It 'throws on non-object input' {
                { Get-TableColumnWidth -InputObject "string" } | Should -Throw
            }

            It 'throws on invalid custom width values' {
                $testObj = [PSCustomObject]@{ Name = 'Test' }
                { Get-TableColumnWidth -InputObject $testObj -CustomWidths @{ Name = 0 } } | Should -Throw
                { Get-TableColumnWidth -InputObject $testObj -CustomWidths @{ Name = 101 } } | Should -Throw
                { Get-TableColumnWidth -InputObject $testObj -CustomWidths @{ Name = "50" } } | Should -Throw
            }

            It 'throws on custom width for non-existent column' {
                $testObj = [PSCustomObject]@{ Name = 'Test' }
                { Get-TableColumnWidth -InputObject $testObj -CustomWidths @{ NonExistent = 50 } } | Should -Throw
            }

            It 'handles case-insensitive property names in custom widths' {
                $testObj = [PSCustomObject]@{ Name = 'Test' }
                $result = Get-TableColumnWidth -InputObject $testObj -CustomWidths @{ name = 50 }
                $result[0] | Should -BeExactly 50
            }
        }
    }
} 
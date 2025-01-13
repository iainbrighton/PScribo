$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module "$moduleRoot\PScribo.psm1" -Force

InModuleScope 'PScribo' {
    # Initialize required variables
    BeforeAll {
        $script:hasCalledTable = $false
        Mock Table { $script:hasCalledTable = $true }
    }

    Describe 'Write-EnhancedTable' {
        BeforeEach {
            $script:hasCalledTable = $false
            $testObject = [PSCustomObject]@{
                Name = "Test"
                Value = "Value"
                Empty = ""
            }
        }

        Context 'Input Validation' {
            It 'throws when InputObject is null' {
                { Write-EnhancedTable -InputObject $null } | Should -Throw
            }

            It 'accepts empty TableParameters hashtable' {
                { Write-EnhancedTable -InputObject $testObject -TableParameters @{} } | Should -Not -Throw
            }

            It 'validates CustomColumnWidths total does not exceed 100' {
                { Write-EnhancedTable -InputObject $testObject -CustomColumnWidths @{ 'Name' = 50; 'Value' = 51 } } | Should -Throw
            }

            It 'validates CustomColumnWidths values are integers between 1 and 100' {
                { Write-EnhancedTable -InputObject $testObject -CustomColumnWidths @{ 'Name' = 0 } } | Should -Throw
                { Write-EnhancedTable -InputObject $testObject -CustomColumnWidths @{ 'Name' = 101 } } | Should -Throw
                { Write-EnhancedTable -InputObject $testObject -CustomColumnWidths @{ 'Name' = 'invalid' } } | Should -Throw
            }
        }

        Context 'By named parameter' {
            It 'creates table with default parameters' {
                Write-EnhancedTable -InputObject $testObject
                $script:hasCalledTable | Should -BeTrue
            }

            It 'passes table parameters correctly' {
                $tableParams = @{ Style = 'Custom' }
                Write-EnhancedTable -InputObject $testObject -TableParameters $tableParams
                $script:hasCalledTable | Should -BeTrue
            }

            It 'removes empty columns when specified' {
                Write-EnhancedTable -InputObject $testObject -RemoveEmptyColumns
                $script:hasCalledTable | Should -BeTrue
            }

            It 'preserves all columns when RemoveEmptyColumns not specified' {
                Write-EnhancedTable -InputObject $testObject
                $script:hasCalledTable | Should -BeTrue
            }
        }

        Context 'PassThru parameter' {
            It 'returns metadata when PassThru specified' {
                $result = Write-EnhancedTable -InputObject $testObject -PassThru
                $result | Should -Not -BeNullOrEmpty
            }

            It 'returns nothing when PassThru not specified' {
                $result = Write-EnhancedTable -InputObject $testObject
                $result | Should -BeNullOrEmpty
            }
        }

        Context 'Pipeline input' {
            It 'accepts pipeline input' {
                { $testObject | Write-EnhancedTable } | Should -Not -Throw
            }

            It 'accumulates all pipeline input before processing' {
                { $testObject, $testObject | Write-EnhancedTable } | Should -Not -Throw
            }
        }

        Context 'Edge Cases' {
            It 'handles single row input' {
                { Write-EnhancedTable -InputObject $testObject } | Should -Not -Throw
            }

            It 'handles objects with different properties' {
                $obj1 = [PSCustomObject]@{ Prop1 = 'Value1' }
                $obj2 = [PSCustomObject]@{ Prop2 = 'Value2' }
                { $obj1, $obj2 | Write-EnhancedTable } | Should -Not -Throw
            }

            It 'handles empty string values correctly' {
                $testObject = [PSCustomObject]@{ Empty = '' }
                Write-EnhancedTable -InputObject $testObject
                $script:hasCalledTable | Should -BeTrue
            }

            It 'handles objects with no properties' {
                $emptyObj = New-Object PSObject
                { Write-EnhancedTable -InputObject $emptyObj } | Should -Not -Throw
            }

            It 'handles mixed property types' {
                $mixedObj = [PSCustomObject]@{
                    String = 'text'
                    Number = 42
                    Date = Get-Date
                    Null = $null
                }
                { Write-EnhancedTable -InputObject $mixedObj } | Should -Not -Throw
            }

            It 'redistributes column widths after removing empty columns' {
                $testObj = [PSCustomObject]@{
                    Name = 'Test'
                    Value = 'Value'
                    Empty1 = ''
                    Empty2 = ''
                }
                $result = Write-EnhancedTable -InputObject $testObj -RemoveEmptyColumns -PassThru
                $result.RemovedColumns | Should -HaveCount 2
                $result.RemovedColumns | Should -Contain 'Empty1'
                $result.RemovedColumns | Should -Contain 'Empty2'
            }
        }
    }
} 
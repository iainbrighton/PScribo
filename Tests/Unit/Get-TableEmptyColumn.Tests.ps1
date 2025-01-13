$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module "$moduleRoot\PScribo.psm1" -Force

InModuleScope 'PScribo' {
    Describe 'Get-TableEmptyColumn' {
        $script:testObject = [PSCustomObject]@{
            Name = "Test"
            Value = "Value"
            Empty = ""
        }

        Context 'Basic functionality' {
            It 'returns array of strings' {
                $result = Get-TableEmptyColumn -InputObject $testObject
                $result.GetType().IsArray -or $result -is [string] | Should Be $true
                if ($result -is [array]) {
                    $result | ForEach-Object { $_.GetType().Name | Should Be 'String' }
                } else {
                    $result.GetType().Name | Should Be 'String'
                }
            }

            It 'identifies empty columns correctly' {
                $result = Get-TableEmptyColumn -InputObject $testObject
                $result | Should Contain 'Empty'
            }

            It 'excludes non-empty columns' {
                $result = Get-TableEmptyColumn -InputObject $testObject
                $result | Should Not Contain 'Name'
                $result | Should Not Contain 'Value'
            }

            It 'accepts pipeline input' {
                $result = $testObject | Get-TableEmptyColumn
                $result | Should Not BeNullOrEmpty
            }

            It 'returns empty array for empty input' {
                $result = Get-TableEmptyColumn -InputObject @()
                $result | Should BeNullOrEmpty
            }

            It 'handles single row input' {
                $singleObj = [PSCustomObject]@{ Empty = ''; NonEmpty = 'value' }
                $result = Get-TableEmptyColumn -InputObject $singleObj
                $result | Should Contain 'Empty'
                $result | Should Not Contain 'NonEmpty'
            }
        }

        Context 'Error handling' {
            It 'throws on null input' {
                { Get-TableEmptyColumn -InputObject $null } | Should Throw
            }

            It 'throws on non-object input' {
                { Get-TableEmptyColumn -InputObject "string" } | Should Throw
            }

            It 'handles objects with no properties' {
                $emptyObj = New-Object PSObject
                $result = Get-TableEmptyColumn -InputObject $emptyObj
                $result | Should BeNullOrEmpty
            }
        }

        Context 'Special cases' {
            It 'handles mixed null and empty string values' {
                $mixedObj = [PSCustomObject]@{
                    Null = $null
                    Empty = ''
                    WhiteSpace = ' '
                    NonEmpty = 'value'
                }
                $result = Get-TableEmptyColumn -InputObject $mixedObj
                $result | Should Contain 'Null'
                $result | Should Contain 'Empty'
                $result | Should Contain 'WhiteSpace'
                $result | Should Not Contain 'NonEmpty'
            }

            It 'handles objects with different properties' {
                $obj1 = [PSCustomObject]@{ Prop1 = ''; Common = '' }
                $obj2 = [PSCustomObject]@{ Prop2 = ''; Common = '' }
                $result = Get-TableEmptyColumn -InputObject @($obj1, $obj2)
                $result | Should Contain 'Prop1'
                $result | Should Contain 'Prop2'
                $result | Should Contain 'Common'
            }

            It 'handles empty objects in array' {
                $obj1 = [PSCustomObject]@{ Prop1 = ''; Common = '' }
                $obj2 = New-Object PSObject
                $result = Get-TableEmptyColumn -InputObject @($obj1, $obj2)
                $result | Should Contain 'Prop1'
                $result | Should Contain 'Common'
            }
        }
    }
} 
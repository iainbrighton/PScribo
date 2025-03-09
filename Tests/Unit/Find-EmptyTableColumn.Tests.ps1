$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module "$moduleRoot\PScribo.psm1" -Force

InModuleScope 'PScribo' {
    Describe 'Find-EmptyTableColumn' {
        $script:testObject = [PSCustomObject]@{
            Name = "Test"
            Value = "Value"
            Empty = ""
        }

        Context 'Basic functionality' {
            It 'returns array of strings' {
                $result = Find-EmptyTableColumn -InputObject $testObject
                $result.GetType().IsArray -or $result -is [string] | Should Be $true
                if ($result -is [array]) {
                    $result | ForEach-Object { $_.GetType().Name | Should Be 'String' }
                } else {
                    $result.GetType().Name | Should Be 'String'
                }
            }

            It 'identifies all empty columns correctly' {
                $result = Find-EmptyTableColumn -InputObject $testObject
                $result | Should Contain 'Empty'
            }

            It 'excludes non-empty columns' {
                $result = Find-EmptyTableColumn -InputObject $testObject
                $result | Should Not Contain 'Name'
                $result | Should Not Contain 'Value'
            }

            It 'returns empty array for empty input' {
                $result = Find-EmptyTableColumn -InputObject @()
                $result | Should BeNullOrEmpty
            }

            It 'handles single row input' {
                $singleObj = [PSCustomObject]@{ Empty = ''; NonEmpty = 'value' }
                $result = Find-EmptyTableColumn -InputObject $singleObj
                $result | Should Contain 'Empty'
                $result | Should Not Contain 'NonEmpty'
            }
        }

        Context 'Whitespace handling' {
            $whitespaceObj = [PSCustomObject]@{
                Space = ' '
                Tab = "`t"
                NewLine = "`n"
                Mixed = " `t`n "
                NonEmpty = 'value'
            }

            It 'treats various whitespace as empty' {
                $result = Find-EmptyTableColumn -InputObject $whitespaceObj
                $result | Should Contain 'Space'
                $result | Should Contain 'Tab'
                $result | Should Contain 'NewLine'
                $result | Should Contain 'Mixed'
            }

            It 'correctly identifies non-empty whitespace-containing values' {
                $obj = [PSCustomObject]@{
                    WithText = ' text '
                    OnlySpace = '   '
                }
                $result = Find-EmptyTableColumn -InputObject $obj
                $result | Should Not Contain 'WithText'
                $result | Should Contain 'OnlySpace'
            }
        }

        Context 'Error handling' {
            It 'throws on null input' {
                { Find-EmptyTableColumn -InputObject $null } | Should Throw
            }

            It 'throws on non-object input' {
                { Find-EmptyTableColumn -InputObject "string" } | Should Throw
            }

            It 'handles objects with no properties' {
                $emptyObj = New-Object PSObject
                $result = Find-EmptyTableColumn -InputObject $emptyObj
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
                $result = Find-EmptyTableColumn -InputObject $mixedObj
                $result | Should Contain 'Null'
                $result | Should Contain 'Empty'
                $result | Should Contain 'WhiteSpace'
                $result | Should Not Contain 'NonEmpty'
            }

            It 'handles objects with different properties' {
                $obj1 = [PSCustomObject]@{ Prop1 = ''; Common = '' }
                $obj2 = [PSCustomObject]@{ Prop2 = ''; Common = '' }
                $result = Find-EmptyTableColumn -InputObject @($obj1, $obj2)
                $result | Should Contain 'Prop1'
                $result | Should Contain 'Prop2'
                $result | Should Contain 'Common'
            }

            It 'handles empty objects in array' {
                $obj1 = [PSCustomObject]@{ Prop1 = ''; Common = '' }
                $obj2 = New-Object PSObject
                $result = Find-EmptyTableColumn -InputObject @($obj1, $obj2)
                $result | Should Contain 'Prop1'
                $result | Should Contain 'Common'
            }
        }
    }
} 
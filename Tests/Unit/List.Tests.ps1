$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    Describe -Name 'List' {

        BeforeEach {

            $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock { }
        }

        It 'returns "PSCustomObject" object' {
            $l = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges'
            }

            $l.Sections[0].GetType().Name | Should -eq 'PSCustomObject'
        }

        It 'creates "PScribo.ListReference" type' {
            $l = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges'
            }

            $l.Sections[0].Type | Should -eq 'PScribo.ListReference'
        }

        It 'creates "PScribo.List" object' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges'
            }

            $pscriboDocument.Lists.Count | Should Be 1
        }

        It 'defaults to "Disc" bulleted list' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges'
            }

            $pscriboDocument.Lists[0].IsNumbered | Should Be $false
            $pscriboDocument.Lists[0].BulletStyle | Should Be 'Disc'
        }

        It 'creates a numbered list' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges' -Numbered
            }

            $pscriboDocument.Lists[0].IsNumbered | Should Be $true
        }

        It 'defaults to "Number" numbered list' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges' -Numbered
            }

            $pscriboDocument.Lists[0].NumberStyle | Should Be 'Number'
            $pscriboDocument.Lists[0].IsNumbered | Should Be $true
        }

        It 'inherits style by default' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges' -Numbered
            }

            $pscriboDocument.Lists[0].IsStyleInherited | Should Be $true
        }

        It 'sets style when specified' {
            $null = Section 'Test' {
                List -Item 'Apples', 'Bananas', 'Oranges' -Style 'Caption'
            }

            $pscriboDocument.Lists[0].Style | Should Be 'Caption'
            $pscriboDocument.Lists[0].IsStyleInherited | Should Be $false
            $pscriboDocument.Lists[0].HasStyle | Should Be $true
        }

        It 'creates multi-level list' {
            $null = Section 'Test' {
                List {
                    Item 'Apples'
                    List -Item 'Jazz', 'Braeburn', 'Pink Lady'
                    Item 'Bananas'
                    Item 'Oranges'
                    List -Item 'Jaffa', 'Satsuma', 'Tangerine'
                }
            }

            $pscriboDocument.Lists[0].IsMultiLevel | Should Be $true
            $pscriboDocument.Lists[0].Items.Count | Should Be 5
            $pscriboDocument.Lists[0].Items[1].Items.Count | Should Be 3
            $pscriboDocument.Lists[0].Items[4].Items.Count | Should Be 3
        }

        It 'creates multi-level numbered and bulleted list' {
            $null = Section 'Test' {
                List -Numbered {
                    Item 'Apples'
                    List -Item 'Jazz', 'Braeburn', 'Pink Lady'
                    Item 'Bananas'
                    Item 'Oranges'
                    List -Item 'Jaffa', 'Satsuma', 'Tangerine'
                }
            }

            $pscriboDocument.Lists[0].IsMultiLevel | Should Be $true
            $pscriboDocument.Lists[0].IsNumbered | Should Be $true
            $pscriboDocument.Lists[0].Items[1].IsNumbered | Should Be $false
            $pscriboDocument.Lists[0].Items[4].IsNumbered | Should Be $false
        }

    }
}

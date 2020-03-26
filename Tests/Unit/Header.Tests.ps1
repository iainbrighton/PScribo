$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Header' {

        $pscriboDocument = Document 'ScaffoldDocument' {}

        Context 'Default Header' {

            It 'returns a PSCustomObject object' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -Default { }
                }
                $pscriboDocument.Header.DefaultHeader -is [PSCustomObject] | Should Be $true
            }

            It 'creates a PScribo.Paragraph type' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -Default { }
                }

                $pscriboDocument.Header.DefaultHeader.Type | Should Be 'PScribo.Header'
            }

            It "also sets 'FirstPageHeader'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -Default -IncludeOnFirstPage { }
                }

                $pscriboDocument.Header.DefaultHeader.Type | Should Be 'PScribo.Header'
                $pscriboDocument.Header.FirstPageHeader.Type | Should Be 'PScribo.Header'
            }

            It "adds additional blank line by default" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -Default { }
                }

                $pscriboDocument.Header.DefaultHeader.Sections.Type | Should Be 'PScribo.BlankLine'
            }

            It "adds no blank line when '-NoSpace' specified" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -Default -NoSpace { }
                }

                $pscriboDocument.Header.DefaultHeader.Sections | Should BeNullOrEmpty
            }

            It "processes embedded 'PScribo.Paragraph'" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -Default -NoSpace {
                        Paragraph 'Test paragraph'
                    }
                }

                $pscriboDocument.Header.DefaultHeader.Sections.Type | Should Be 'PScribo.Paragraph'
            }

            It "processes embedded 'PScribo.Table'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -Default -NoSpace {
                        Table -Hashtable ([Ordered] @{ 'Left' = 'Right' }) -List
                    }
                }

                $pscriboDocument.Header.DefaultHeader.Sections.Type | Should Be 'PScribo.Table'
            }
        }

        Context 'First Page Header' {

            It 'returns a PSCustomObject object' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -FirstPage { }
                }
                $pscriboDocument.Header.FirstPageHeader -is [PSCustomObject] | Should Be $true
            }

            It 'creates a PScribo.Paragraph type' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -FirstPage { }
                }

                $pscriboDocument.Header.FirstPageHeader.Type | Should Be 'PScribo.Header'
            }

            It "adds additional blank line by default" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -FirstPage { }
                }

                $pscriboDocument.Header.FirstPageHeader.Sections.Type | Should Be 'PScribo.BlankLine'
            }

            It "adds no blank line when '-NoSpace' specified" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -FirstPage -NoSpace { }
                }

                $pscriboDocument.Header.FirstPageHeader.Sections | Should BeNullOrEmpty
            }

            It "processes embedded 'PScribo.Paragraph'" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Header -FirstPage -NoSpace {
                        Paragraph 'Test paragraph'
                    }
                }

                $pscriboDocument.Header.FirstPageHeader.Sections.Type | Should Be 'PScribo.Paragraph'
            }

            It "processes embedded 'PScribo.Table'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Header -FirstPage -NoSpace {
                        Table -Hashtable ([Ordered] @{ 'Left' = 'Right' }) -List
                    }
                }

                $pscriboDocument.Header.FirstPageHeader.Sections.Type | Should Be 'PScribo.Table'
            }
        }
    }
}

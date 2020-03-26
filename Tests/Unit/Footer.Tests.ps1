$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Footer' {

        $pscriboDocument = Document 'ScaffoldDocument' {}

        Context 'Default Footer' {

            It 'returns a PSCustomObject object' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -Default { }
                }
                $pscriboDocument.Footer.DefaultFooter -is [PSCustomObject] | Should Be $true
            }

            It 'creates a PScribo.Paragraph type' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -Default { }
                }

                $pscriboDocument.Footer.DefaultFooter.Type | Should Be 'PScribo.Footer'
            }

            It "also sets 'FirstPageFooter'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -Default -IncludeOnFirstPage { }
                }

                $pscriboDocument.Footer.DefaultFooter.Type | Should Be 'PScribo.Footer'
                $pscriboDocument.Footer.FirstPageFooter.Type | Should Be 'PScribo.Footer'
            }

            It "adds additional blank line by default" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -Default { }
                }

                $pscriboDocument.Footer.DefaultFooter.Sections.Type | Should Be 'PScribo.BlankLine'
            }

            It "adds no blank line when '-NoSpace' specified" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -Default -NoSpace { }
                }

                $pscriboDocument.Footer.DefaultFooter.Sections | Should BeNullOrEmpty
            }

            It "processes embedded 'PScribo.Paragraph'" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -Default -NoSpace {
                        Paragraph 'Test paragraph'
                    }
                }

                $pscriboDocument.Footer.DefaultFooter.Sections.Type | Should Be 'PScribo.Paragraph'
            }

            It "processes embedded 'PScribo.Table'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -Default -NoSpace {
                        Table -Hashtable ([Ordered] @{ 'Left' = 'Right' }) -List
                    }
                }

                $pscriboDocument.Footer.DefaultFooter.Sections.Type | Should Be 'PScribo.Table'
            }
        }

        Context 'First Page Footer' {

            It 'returns a PSCustomObject object' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -FirstPage { }
                }
                $pscriboDocument.Footer.FirstPageFooter -is [PSCustomObject] | Should Be $true
            }

            It 'creates a PScribo.Paragraph type' {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -FirstPage { }
                }

                $pscriboDocument.Footer.FirstPageFooter.Type | Should Be 'PScribo.Footer'
            }

            It "adds additional blank line by default" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -FirstPage { }
                }

                $pscriboDocument.Footer.FirstPageFooter.Sections.Type | Should Be 'PScribo.BlankLine'
            }

            It "adds no blank line when '-NoSpace' specified" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -FirstPage -NoSpace { }
                }

                $pscriboDocument.Footer.FirstPageFooter.Sections | Should BeNullOrEmpty
            }

            It "processes embedded 'PScribo.Paragraph'" {
                $pscriboDocument =  Document 'ScaffoldDocument' {
                    Footer -FirstPage -NoSpace {
                        Paragraph 'Test paragraph'
                    }
                }

                $pscriboDocument.Footer.FirstPageFooter.Sections.Type | Should Be 'PScribo.Paragraph'
            }

            It "processes embedded 'PScribo.Table'" {
                $pscriboDocument = Document 'ScaffoldDocument' {
                    Footer -FirstPage -NoSpace {
                        Table -Hashtable ([Ordered] @{ 'Left' = 'Right' }) -List
                    }
                }

                $pscriboDocument.Footer.FirstPageFooter.Sections.Type | Should Be 'PScribo.Table'
            }
        }
    }
}

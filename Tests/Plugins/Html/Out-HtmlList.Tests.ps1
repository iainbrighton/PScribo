$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlList' {

        BeforeEach {

            ## Scaffold document options
            $Document = Document -Name 'TestDocument' -ScriptBlock {}
            # $script:currentPageNumber = 1
            $pscriboDocument = $Document
            $Options = New-PScriboTextOption
        }

        It 'Single-level bulleted disc list' {

            $bulletStyle = 'Disc'
            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle $bulletStyle }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match ('<ul style="list-style-type:{0};">' -f $bulletStyle)
            $result | Should Match ('<li>{0}</li>' -f $testItems[0])
            $result | Should Match ('<li>{0}</li>' -f $testItems[1])
            $result | Should Match ('<li>{0}</li>' -f $testItems[2])
            $result | Should Match '</ul>'
        }

        It 'Single-level bulleted circle list' {

            $bulletStyle = 'Circle'
            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle $bulletStyle }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match ('<ul style="list-style-type:{0};">' -f $bulletStyle)
            $result | Should Match ('<li>{0}</li>' -f $testItems[0])
            $result | Should Match ('<li>{0}</li>' -f $testItems[1])
            $result | Should Match ('<li>{0}</li>' -f $testItems[2])
            $result | Should Match '</ul>'
        }

        It 'Single-level bulleted dash list' {

            $bulletStyle = 'Dash'
            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle $bulletStyle }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match '<ul>' # Dash is default unordered list type
            $result | Should Match ('<li>{0}</li>' -f $testItems[0])
            $result | Should Match ('<li>{0}</li>' -f $testItems[1])
            $result | Should Match ('<li>{0}</li>' -f $testItems[2])
            $result | Should Match '</ul>'
        }

        It 'Single-level bulleted square list' {

            $bulletStyle = 'Square'
            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle $bulletStyle }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match ('<ul style="list-style-type:{0};">' -f $bulletStyle)
            $result | Should Match ('<li>{0}</li>' -f $testItems[0])
            $result | Should Match ('<li>{0}</li>' -f $testItems[1])
            $result | Should Match ('<li>{0}</li>' -f $testItems[2])
            $result | Should Match '</ul>'
        }

        It 'Single-level numbered list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -Numbered }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match '<ol style="list-style-type:decimal;">'
            $result | Should Match ('<li>{0}</li>' -f $testItems[0])
            $result | Should Match ('<li>{0}</li>' -f $testItems[1])
            $result | Should Match ('<li>{0}</li>' -f $testItems[2])
            $result | Should Match '</ol>'
        }

        It 'Multi-level numbered list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $testSubItems = 'Braeburn', 'Granny Smith'
            $null = Section 'Test' {
                List -Numbered {
                    Item $testItems[0]
                    List -Numbered {
                        Item $testSubItems[0]
                        Item $testSubItems[1]
                    }
                    Item $testItems[1]
                    Item $testItems[2]
                }
            }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match '<ol style="list-style-type:decimal;">'
            $result | Should Match ('<li>{0}</li>\s+<ol style="list-style-type:decimal;">\s+<li>{1}</li>\s+<li>{2}</li>\s+</ol>' -f $testItems[0], $testSubItems[0], $testSubItems[1])
        }

        It 'Multi-level roman numbered list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $testSubItems = 'Braeburn', 'Granny Smith'
            $null = Section 'Test' {
                List -Numbered -NumberStyle Roman {
                    Item $testItems[0]
                    List -Numbered {
                        Item $testSubItems[0]
                        Item $testSubItems[1]
                    }
                    Item $testItems[1]
                    Item $testItems[2]
                }
            }

            $result = Out-HtmlList -List $Document.Lists[0]

            $result | Should Match '<ol style="list-style-type:lower-roman;">'
            $result | Should Match ('<li>{0}</li>\s+<ol style="list-style-type:lower-roman;">\s+<li>{1}</li>\s+<li>{2}</li>\s+</ol>' -f $testItems[0], $testSubItems[0], $testSubItems[1])
        }

    }
}

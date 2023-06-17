$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextList' {

        BeforeEach {

            ## Scaffold document options
            $Document = Document -Name 'TestDocument' -ScriptBlock {}
            # $script:currentPageNumber = 1
            $pscriboDocument = $Document
            $Options = New-PScriboTextOption
        }

        It 'Terminates with a blank line' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems }

            $result = Out-TextList -List $Document.Lists[0]

            $expectedMatch = '{0}{0}$' -f [System.Environment]::NewLine
            $result | Should Match $expectedMatch
        }

        It 'Single-level bulleted disc list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle Disc }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('\* {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('\* {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('\* {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

        It 'Single-level bulleted circle list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle Circle }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('o {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('o {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('o {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

        It 'Single-level bulleted dash list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle Dash }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('- {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('- {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('- {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

        It 'Single-level bulleted square list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -BulletStyle Square }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('\* {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('\* {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('\* {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

        It 'Single-level numbered list' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $null = Section 'Test' { List -Item $testItems -Numbered }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('1\. {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('2\. {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('3\. {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
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

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('1\. {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('1\. {0}{1}' -f $testSubItems[0], [System.Environment]::NewLine)
            $result | Should Match ('2\. {0}{1}' -f $testSubItems[1], [System.Environment]::NewLine)
            $result | Should Match ('2\. {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('3\. {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

        It 'outputs custom numbered list format' {

            $testItems = 'Apples', 'Bananas', 'Oranges'
            $customNumberFormat = 'xYz-%%%.'
            $indent = 1500
            $hanging = 200

            $null = Section 'Test' {
                NumberStyle -Id 'CustomNumberStyle' -Custom $customNumberFormat -Indent $indent -Hanging $hanging -Align Left
                List -Numbered -NumberStyle CustomNumberStyle -Item $testItems
            }

            $result = Out-TextList -List $Document.Lists[0]

            $result | Should Match ('xYz-001. {0}{1}' -f $testItems[0], [System.Environment]::NewLine)
            $result | Should Match ('xYz-002. {0}{1}' -f $testItems[1], [System.Environment]::NewLine)
            $result | Should Match ('xYz-003. {0}{1}' -f $testItems[2], [System.Environment]::NewLine)
        }

    }
}

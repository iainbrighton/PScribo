$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\ConvertTo-AlignedString' {

        It 'Defaults to left alignment and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '{0}{1}' -f $testString, [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50

            $result | Should Be $expected
        }

        It 'Aligns to left alignment without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = $testString

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50 -NoNewLine

            $result | Should Be $expected
        }

        It 'Aligns to left alignment and indents with tabs' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '        {0}{1}' -f $testString, [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Tabs 2

            $result | Should Be $expected
        }

        It 'Pads string when using right alignment and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '     {0}{1}' -f $testString, [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50 -Align Right

            $result | Should Be $expected
        }

        It 'Pads string when using right alignment without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '     {0}' -f $testString

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50 -Align Right -NoNewLine

            $result | Should Be $expected
        }

        It 'Pads string when using center alignment and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '  {0}{1}' -f $testString, [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50 -Align Center

            $result | Should Be $expected
        }

        It 'Pads string when using center alignment without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '  {0}' -f $testString

            $result = ConvertTo-AlignedString -InputObject $testString -Width 50 -Align Center -NoNewLine

            $result | Should Be $expected
        }

        It 'Wraps text at empty space when over the specified width and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = 'The quick brown fox{0}jumped over the lazy dog.{0}' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25

            $result | Should Be $expected
        }

        It 'Wraps text at empty space when over the specified width without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = 'The quick brown fox{0}jumped over the lazy dog.' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25 -NoNewLine

            $result | Should Be $expected
        }

        It 'Wraps unbroken line when over the specified width' {
            $testString = 'The-quick-brown-fox-jumped-over-the-lazy-dog.'
            $expected = 'The-quick-brown-fox-jumpe{0}d-over-the-lazy-dog.{0}' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25

            $result | Should Be $expected
        }

        It 'Wraps unbroken line when over the specified width without new line' {
            $testString = 'The-quick-brown-fox-jumped-over-the-lazy-dog.'
            $expected = 'The-quick-brown-fox-jumpe{0}d-over-the-lazy-dog.' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25 -NoNewLine

            $result | Should Be $expected
        }

        It 'Wraps text at blank line and indents with tabs and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '        The quick brown{0}        fox jumped over{0}        the lazy dog.{0}' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25 -Tabs 2

            $result | Should Be $expected
        }

        It 'Wraps text at blank line and indents with tabs without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '        The quick brown{0}        fox jumped over{0}        the lazy dog.' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 25 -Tabs 2 -NoNewLine

            $result | Should Be $expected
        }

        It 'Pads string when using center alignment with tabs and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '            {0}{1}' -f $testString, [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 60 -Align Center -Tabs 2

            $result | Should Be $expected
        }

        It 'Wraps text at empty space and centers and adds new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '  The quick brown fox{0} jumped over the lazy{0}         dog.{0}' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 22 -Align Center

            $result | Should Be $expected
        }

        It 'Wraps text at empty space and centers without new line' {
            $testString = 'The quick brown fox jumped over the lazy dog.'
            $expected = '  The quick brown fox{0} jumped over the lazy{0}         dog.' -f [System.Environment]::NewLine

            $result = ConvertTo-AlignedString -InputObject $testString -Width 22 -Align Center -NoNewLine

            $result | Should Be $expected
        }
    }
}

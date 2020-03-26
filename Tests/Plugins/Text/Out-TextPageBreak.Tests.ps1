$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    $isNix = $false
    if (($PSVersionTable['PSEdition'] -eq 'Core') -and (-not $IsWindows))
    {
        $isNix = $true
    }


    Describe 'Plugins\Text\Out-TextPageBreak' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $Options = New-PScriboTextOption

        It 'Defaults to 120 and includes 2 new lines' {
            $expected = 124
            if ($isNix) { $expected -= 2 }

            $l = Out-TextPageBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes 2 new lines' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 40
            $expected = 44
            if ($isNix) { $expected -= 2 }

            $l = Out-TextPageBreak

            $l.Length | Should Be $expected
        }

        It 'Wraps lines and includes new line' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 80
            $expected = 86
            if ($isNix) { $expected -= 3 }

            $l = Out-TextPageBreak

            $l.Length | Should Be $expected
        }
    }
}

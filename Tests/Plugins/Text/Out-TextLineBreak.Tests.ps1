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


    Describe 'Plugins\Text\Out-TextLineBreak' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {}
        $Options = New-PScriboTextOption;

        It 'Defaults to 120 and includes new line' {
            $expected = 122
            if ($isNix) { $expected -= 1 }

            $l = Out-TextLineBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes new line' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 40
            $expected = 42
            if ($isNix) { $expected -= 1 }

            $l = Out-TextLineBreak

            $l.Length | Should Be $expected
        }

        It 'Wraps lines and includes new line' {
            $Options = New-PScriboTextOption -TextWidth 40 -SeparatorWidth 80
            $expected = 84
            if ($isNix) { $expected -= 2 }

            $l = Out-TextLineBreak

            $l.Length | Should Be $expected
        }
    }
}

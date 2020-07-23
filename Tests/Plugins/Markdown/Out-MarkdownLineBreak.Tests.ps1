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

    Describe 'Plugins\Markdown\Out-MarkdownLineBreak' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {}
        $Options = New-PScriboMarkdownOption

        It 'Defaults to 10 and includes 2 new lines' {
            $expected = 14
            if ($isNix) { $expected -= 2 }

            $l = Out-MarkdownLineBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes 2 new lines' {
            $Options = New-PScriboMarkdownOption -TextWidth 40 -LineBreakSeparatorWidth 80
            $expected = 44
            if ($isNix) { $expected -= 2 }

            $l = Out-MarkdownLineBreak

            $l.Length | Should Be $expected
        }
    }
}

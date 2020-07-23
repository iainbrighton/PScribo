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

    Describe 'Plugins\Markdown\Out-MarkdownPageBreak' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $Options = New-PScriboMarkdownOption

        It 'Defaults to 20 and includes 2 new lines' {
            $expected = 24
            if ($isNix) { $expected -= 2 }

            $l = Out-MarkdownPageBreak

            $l.Length | Should Be $expected
        }

        It 'Truncates to 40 and includes 2 new lines' {
            $Options = New-PScriboMarkdownOption -TextWidth 40 -PageBreakSeparatorWidth 80
            $expected = 44
            if ($isNix) { $expected -= 2 }

            $l = Out-MarkdownPageBreak

            $l.Length | Should Be $expected
        }
    }
}

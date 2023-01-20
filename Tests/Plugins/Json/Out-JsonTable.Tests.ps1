$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Json\Out-JsonTable' {

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $pscriboDocument = $Document

        ## Context (list vs table) doesn't make a difference as they are treated the same
        
        $services = @(
            [Ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
            [Ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
            [Ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
        )

        It 'Input matches output' {
            $expected = 3

            $table = Table -Hashtable $services 'Test Table' | Out-JsonTable

            $table.Count | Should Be $expected
        }

        It 'Caption present' {
            $expected = $true

            $table = Table -Hashtable $services 'Test Table' -Caption 'Test' | Out-JsonTable

            $table[0].Caption | Should Be $expected
        }

        It 'Input matches output with caption present' {
            $expected = 3

            $table = Table -Hashtable $services 'Test Table' -Caption 'Test' | Out-JsonTable

            $table[1].Count | Should Be $expected
        }
    }
}

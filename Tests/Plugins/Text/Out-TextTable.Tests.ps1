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

    Describe 'Plugins\Text\Out-TextTable' {

        ## Scaffold document options
        $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {}

        Context 'As Table' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120' {
                $expected = 208
                if ($isNix) { $expected -= 6 }

                $table = Table -Hashtable $services -Name 'Test Table' | Out-TextTable

                $table.Length | Should Be $expected # Trailing spaces are removed (#67)
            }

            It 'Set width with of 35' {
                $Options = New-PScriboTextOption -TextWidth 35
                $expected = 313
                if ($isNix) { $expected -= 10 }

                $table = Table -Hashtable $services -Name 'Test Table' | Out-TextTable

                $table.Length | Should Be $expected ## Text tables are now set to wrap.. Trailing spaces are removed (#67)
            }
        }

        Context 'As List' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120' {
                $expected = 255
                if ($isNix) { $expected -= 12 }

                $table = Table -Hashtable $services 'Test Table' -List | Out-TextTable

                $table.Length | Should Be $expected
            }

            It 'Default width of 25' {
                $Options = New-PScriboTextOption -TextWidth 25
                $expected = 354
                if ($isNix) { $expected -= 18 }

                $table = Table -Hashtable $services 'Test Table' -List | Out-TextTable

                $table.Length | Should Be $expected # Trailing spaces are removed (#67)
            }
        }
    }
}

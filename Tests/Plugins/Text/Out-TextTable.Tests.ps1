$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextTable' {

        $isNix = (($PSVersionTable['PSEdition'] -eq 'Core') -and (-not $IsWindows))

        ## Scaffold document options
        $Document = Document -Name 'TestDocument' -ScriptBlock {}
        $script:currentPageNumber = 1
        $pscriboDocument = $Document

        Context 'As Table' {

            $services = @(
                [Ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [Ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [Ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120' {
                $expected = 206
                if ($isNix) { $expected -= 5 } # Account for line ending differences between Windows and *nix

                $table = Table -Hashtable $services -Name 'Test Table' | Out-TextTable

                $table.Length | Should Be $expected
            }

            It 'Set width with of 35' {

                if ($isNix)
                {
                    Set-ItResult -Inconclusive -Because "of inconsistent Window/*nix 'Format-Table' output text wrapping"
                }
                else
                {
                    $Options = New-PScriboTextOption -TextWidth 35
                    $expected = 311

                    $table = Table -Hashtable $services -Name 'Test Table' | Out-TextTable

                    $table.Length | Should Be $expected ## Text tables are now set to wrap.. Trailing spaces are removed (#67)
                }
            }
        }

        Context 'As List' {

            $services = @(
                [Ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [Ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [Ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'Default width of 120' {
                $expected = 253
                if ($isNix) { $expected -= 11 } # Account for line ending differences between Windows and *nix

                $table = Table -Hashtable $services 'Test Table' -List | Out-TextTable

                $table.Length | Should Be $expected
            }

            It 'Default width of 25' {

                if ($isNix)
                {
                    Set-ItResult -Inconclusive -Because "of inconsistent Window/*nix 'Format-List' output text wrapping"
                }
                else
                {
                    $Options = New-PScriboTextOption -TextWidth 25
                    $expected = 352

                    $table = Table -Hashtable $services 'Test Table' -List | Out-TextTable

                    $table.Length | Should Be $expected
                }

            }
        }
    }
}

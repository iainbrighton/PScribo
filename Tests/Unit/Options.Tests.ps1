$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Options\Merge-PScriboPluginOption' {

        It 'does not throw with empty hashtable(s)' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ PageOrientation = 'Portrait' }
                DefaultPluginOptions = @{ }
                PluginOptions = @{ }
            }

            { Merge-PScriboPluginOption @mergePScriboPluginOptionParams } | Should Not Throw;
        }

        It 'does not throw with $null hashtable(s)' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ PageOrientation = 'Portrait' }
                DefaultPluginOptions = $null;
                PluginOptions = $null;
            }

            { Merge-PScriboPluginOption @mergePScriboPluginOptionParams } | Should Not Throw;
        }

        It 'merges document options, default plugin and specified plugin options' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ DocumentOption = 'Document' }
                DefaultPluginOptions = @{ DefaultPluginOption = 'DefaultPlugin' }
                PluginOptions = @{ PluginOption = 'Plugin' }
            }

            $result =  Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

            $result.Keys.Count | Should Be 3;
        }

        It 'returns document options by default' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
            }

            $result =  Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'Document';
        }

        It 'overwrites document options with default plugin options' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
                DefaultPluginOptions = @{ TestOption = 'DefaultPlugin' }
            }

            $result =  Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'DefaultPlugin';
        }

        It 'overwrites document options and default plugin options with specified plugin options' {
            $mergePScriboPluginOptionParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
                DefaultPluginOptions = @{ TestOption = 'DefaultPlugin' }
                PluginOptions = @{ TestOption = 'Plugin' }
            }

            $result =  Merge-PScriboPluginOption @mergePScriboPluginOptionParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'Plugin';
        }

    } #end describe

} #end inmodulescope

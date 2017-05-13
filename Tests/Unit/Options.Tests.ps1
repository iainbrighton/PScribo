$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Options\Merge-PScriboPluginOptions' {

        It 'does not throw with empty hashtable(s)' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ PageOrientation = 'Portrait' }
                DefaultPluginOptions = @{ }
                PluginOptions = @{ }
            }

            { Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams } | Should Not Throw;
        }

        It 'does not throw with $null hashtable(s)' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ PageOrientation = 'Portrait' }
                DefaultPluginOptions = $null;
                PluginOptions = $null;
            }

            { Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams } | Should Not Throw;
        }

        It 'merges document options, default plugin and specified plugin options' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ DocumentOption = 'Document' }
                DefaultPluginOptions = @{ DefaultPluginOption = 'DefaultPlugin' }
                PluginOptions = @{ PluginOption = 'Plugin' }
            }

            $result =  Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams;

            $result.Keys.Count | Should Be 3;
        }

        It 'returns document options by default' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
            }

            $result =  Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'Document';
        }

        It 'overwrites document options with default plugin options' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
                DefaultPluginOptions = @{ TestOption = 'DefaultPlugin' }
            }

            $result =  Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'DefaultPlugin';
        }

        It 'overwrites document options and default plugin options with specified plugin options' {
            $mergePScriboPluginOptionsParams = @{
                DocumentOptions = @{ TestOption = 'Document' }
                DefaultPluginOptions = @{ TestOption = 'DefaultPlugin' }
                PluginOptions = @{ TestOption = 'Plugin' }
            }

            $result =  Merge-PScriboPluginOptions @mergePScriboPluginOptionsParams;

            $result.Keys.Count | Should Be 1;
            $result['TestOption'] | Should Be 'Plugin';
        }

    } #end describe

} #end inmodulescope

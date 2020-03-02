$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Text\Out-TextImage' {

        It 'adds alttext to image' {
            $testImage = [PSCustomObject] @{
                Text     = 'Dummy Image'
            }
            $expected = '\[Image Text="{0}"\]' -f $testImage.Text

            $result = Out-TextImage -Image $testImage

            $result -match $expected | Should Be $true
        }
    }
}

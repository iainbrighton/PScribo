$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    $testRoot = Split-Path -Path $PSScriptRoot -Parent

    Describe -Name 'Image' -Fixture {

        $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock { }
        $testJpgFile = Join-Path -Path $testRoot -ChildPath 'TestImage.jpg'
        $testPngFile = Join-Path -Path $testRoot -ChildPath 'TestImage.png'

        It -name 'returns "PSCustomObject" object.' -test {
            $p = Image -Path $testJpgFile
            $p.GetType().Name | Should -eq 'PSCustomObject'
        }

        It -name 'creates "PScribo.Image" type.' -test {
            $p = Image -Path $testJpgFile
            $p.Type | Should -eq 'PScribo.Image'
        }

        It -name 'set MIME type "jpeg"' -test {
            $p = Image -Path $testJpgFile
            $p.MIMEType | Should -eq 'image/jpeg'
        }

        It -name 'set MIME type "png"' -test {
            $p = Image -Path $testPngFile
            $p.MIMEType | Should -eq 'image/png'
        }

        It -name 'creates image number' {
            $p = Image -Path $testJpgFile
            $p.Name | Should -match '^Image\d+$'
        }

        It -name 'sets Emu width' -test {
            $p = Image -Path $testJpgFile -Height 61 -Width 250
            $p.WidthEm | Should -eq 2381250
        }

        It -name 'sets Emu Height' -test {
            $p = Image -Path $testJpgFile -Height 61 -Width 250
            $p.HeightEm | Should -eq 581025
        }
    }
}

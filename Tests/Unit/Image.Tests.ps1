$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

InModuleScope -ModuleName 'PScribo' -ScriptBlock {

    $testRoot = Split-Path -Path $PSScriptRoot -Parent;

    Describe -Name 'Image\Image' -Fixture {

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
    } #end describe Image\Image

    Describe -Name 'Image\Resolve-ImageUri' -Fixture {

        It "returns 'System.Uri' type" {

            $testPath = 'about:Blank'

            $result = Resolve-ImageUri -Path $testPath

            $result -is [System.Uri] | Should Be $true
        }

    } #end describe Image\Resolve-ImageUri

    Describe -Name 'Image\Get-UriBytes' -Fixture {

        It "returns 'System.Byte[]' type" {

            $testPath = Join-Path -Path $testRoot -ChildPath 'TestImage.jpg'
            $testUri = Resolve-ImageUri -Path $testPath

            $result = Get-UriBytes -Uri $testUri

            $result -is [System.Object[]] | Should Be $true
            $result[0] -is [System.Byte] | Should Be $true
        }

    } #end describe Image\Get-UriBytes

    Describe -Name 'Image\ConvertTo-Image' -Fixture {

        It "returns 'System.Drawing.Image' type" {

            $testPath = Join-Path -Path $testRoot -ChildPath 'TestImage.jpg'
            $testUri = Resolve-ImageUri -Path $testPath
            $testBytes = Get-UriBytes -Uri $testUri

            $result = ConvertTo-Image -Bytes $testBytes

            $result -is [System.Drawing.Image] | Should Be $true
        }

    } #end describe Image\ConvertTo-Image

    Describe -Name 'Image\Get-PScriboImage' -Fixture {


        $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock {
            Image -Path (Join-Path -Path $testRoot -ChildPath 'TestImage.jpg') -Id 1
            Section Nested {
                Image -Path (Join-Path -Path $testRoot -ChildPath 'TestImage.png') -Id 2
            }
        }

        It 'finds all Images' {

            $result = Get-PScriboImage -Section $pscriboDocument.Sections

            $result.Count | Should Be 2
        }

        It 'finds Image by Id' {

            $result = @(Get-PScriboImage -Section $pscriboDocument.Sections -Id 2)

            $result.Count | Should Be 1
        }

    } #end describe Image\Get-PScriboImage

}

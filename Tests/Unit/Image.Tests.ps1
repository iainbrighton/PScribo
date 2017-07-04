$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$testRoot  = Split-Path -Path $here -Parent
$moduleRoot = Split-Path -Path $testRoot -Parent
Import-Module -Name "$moduleRoot\PScribo.psm1" -Force

Function New-TestImage 
{
    param(
        $fileName
    )
    Add-Type -AssemblyName System.Drawing 
    $bmp = New-Object -TypeName System.Drawing.Bitmap -ArgumentList 250, 61 
    $font = New-Object -TypeName System.Drawing.Font -ArgumentList Consolas, 24 
    $brushBg = [System.Drawing.Brushes]::Yellow 
    $brushFg = [System.Drawing.Brushes]::Black 
    $graphics = [System.Drawing.Graphics]::FromImage($bmp) 
    $graphics.FillRectangle($brushBg,0,0,$bmp.Width,$bmp.Height) 
    $graphics.DrawString('Hello World',$font,$brushFg,10,10) 
    $graphics.Dispose() 
    $bmp.Save($fileName) 
}

InModuleScope -ModuleName 'PScribo' -ScriptBlock {
    Describe -Name 'Image' -Fixture {
        $pscriboDocument = Document -Name 'ScaffoldDocument' -ScriptBlock {}
        $TestFile = "$TestDrive\Test.jpg"
        #Build TestFile
        $null = New-TestImage -fileName $TestFile
        #
        It -name 'returns a PSCustomObject object.' -test {
            $p = Image -FilePath $TestFile
            $p.GetType().Name | Should -LegacyArg1 Be -LegacyArg2 'PSCustomObject'
        }
        It -name 'creates a PScribo.Image type.' -test {
            $p = Image -FilePath $TestFile
            $p.Type | Should -LegacyArg1 Be -LegacyArg2 'PScribo.Image'
        }
        It -name 'MIME image/jpeg' -test {
            $p = Image -FilePath $TestFile
            $p.MIME | Should -LegacyArg1 Be -LegacyArg2 'image/jpeg'
        }
        It 'Name should be Img3.jpg' {
            $p = Image -FilePath $TestFile
            $p.Name | Should -LegacyArg1 Be -LegacyArg2 'Img3.jpg'
        }
        It -name 'PixelHeight Without Pixel Parameters' -test {
            $p = Image -FilePath $TestFile
            $p.PixelHeight | Should -LegacyArg1 Be -LegacyArg2 '61'
        }
        It -name 'PixelWidth Without Pixel Parameters' -test {
            $p = Image -FilePath $TestFile
            $p.PixelWidth | Should -LegacyArg1 Be -LegacyArg2 '250'
        }
        It -name 'Emu width' -test {
            $p = Image -FilePath $TestFile
            $p.EMUWidth | Should -LegacyArg1 Be -LegacyArg2 2381250
        }
        It -name 'EMU Height' -test {
            $p = Image -FilePath $TestFile
            $p.EMUHeight | Should -LegacyArg1 Be -LegacyArg2 581025
        }
        It -name 'With Pixel Height Parameters' -test {
            $p = Image -FilePath $TestFile -PixelHeight 100
            $p.PixelHeight | Should -LegacyArg1 Be -LegacyArg2 100
        }
        It -name 'With Pixel Width Parameters' -test {
            $p = Image -FilePath $TestFile -PixelWidth 100
            $p.PixelWidth | Should -LegacyArg1 Be -LegacyArg2 100
        }

    }
} #end describe Image

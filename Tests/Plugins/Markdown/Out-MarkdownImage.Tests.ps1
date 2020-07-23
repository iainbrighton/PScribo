$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Markdown\Out-MarkdownImage' {

        It 'adds alttext to image' {
            $testImage = [PSCustomObject] @{
                Id          = '0'
                ImageNumber = 0
                Text        = 'Test Image'
                Type        = 'PScribo.Image'
                Bytes       = [byte[]] @(0,1,2,3)
                Uri         = 'https://non.existent/testimage.png'
                Name        = 'Image0'
                Align       = 'Left'
                MIMEType    = 'image/png'
                WidthEm     = 1.00
                HeightEm    = 1.00
                Width       = 16
                Height      = 16
            }

            $result = Out-MarkdownImage -Image $testImage

            $result | Should Match ('!\[{0}\]\({1}\){2}{2}' -f $testImage.Text, $testImage.Uri, [System.Environment]::NewLine)
        }

        It 'outputs Base64 encoded file at bottom of document when "EmbedImage" is specified' {

            $testDocument = Document -Name 'TestDocument' -ScriptBlock {
                Image -Text 'Test' -Uri 'https://cdn.pixabay.com/photo/2014/08/26/19/20/document-428334_640.jpg'
            }

            $result = Get-MarkdownDocument -Document $testDocument -Options (New-PScriboMarkdownOption -EmbedImage $true)

            $result | Should Match '\[ref_image0\]: data:image/jpeg;base64,'
        }
    }
}

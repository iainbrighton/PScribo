$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Style\Style' {
        $pscriboDocument = Document 'ScaffoldDocument' {};
        $pscriboDocumentStyleCount = $pscriboDocument.Styles.Count;

        Context 'By Named Parameter' {

            It 'sets style using named -Id and -Font parameters.' {
                Style -Name 'Test Style' -Font Arial;
                $pscriboDocument.Styles['TestStyle'].Name | Should BeExactly 'Test Style';
                $pscriboDocument.Styles['TestStyle'].Id | Should BeExactly 'TestStyle';
            }

            It 'sets style font size by named -Size parameter.' {
                Style -Name 'Test Style' -Font Arial -Size 14;
                $pscriboDocument.Styles['TestStyle'].Size | Should Be 14;
            }

            It 'sets style color to white (ffffff) by named -Color parameter.' {
                Style -Name 'Test Style' -Font Arial -Color FFFFFF;
                $pscriboDocument.Styles['TestStyle'].Color | Should BeExactly 'ffffff';
            }

            It 'sets style color to red (ff0000) by named -Colour parameter.' {
                Style -Name 'Test Style' -Font Arial -Colour FF0000;
                $pscriboDocument.Styles['TestStyle'].Color | Should BeExactly 'ff0000';
            }

            It 'sets style background color to blue (0000ff) by named -Color parameter.' {
                Style -Name 'Test Style' -Font Arial -BackgroundColor 0000FF;
                $pscriboDocument.Styles['TestStyle'].BackgroundColor | Should BeExactly '0000ff';
            }

            It 'sets style background color to green (00ff00) by named -Colour parameter.' {
                Style -Name 'Test Style' -Font Arial -BackgroundColour 00FF00;
                $pscriboDocument.Styles['TestStyle'].BackgroundColor | Should BeExactly '00ff00';
            }

            It 'sets style font to bold.' {
                $pscriboDocument.Styles['TestStyle'].Bold | Should Be $false;
                Style -Name 'Test Style' -Font Arial -Bold;
                $pscriboDocument.Styles['TestStyle'].Bold | Should Be $true;
            }

            It 'sets style font to italic.' {
                $pscriboDocument.Styles['TestStyle'].Italic | Should Be $false;
                Style -Name 'Test Style' -Font Arial -Italic;
                $pscriboDocument.Styles['TestStyle'].Italic | Should Be $true;
            }

            It 'sets style font to underline.' {
                $pscriboDocument.Styles['TestStyle'].Underline | Should Be $false;
                Style -Name 'Test Style' -Font Arial -Underline;
                $pscriboDocument.Styles['TestStyle'].Underline | Should Be $true;
            }

            It 'sets text alignment to center.' {
                Style -Name 'Test Style' -Font Arial -Align Center;
                $pscriboDocument.Styles['TestStyle'].Align | Should Be 'Center';
            }

            It 'sets text alignment to right.' {
                Style -Name 'Test Style' -Font Arial -Align Right;
                $pscriboDocument.Styles['TestStyle'].Align | Should Be 'Right';
            }

            It 'sets text alignment to justified.' {
                Style -Name 'Test Style' -Font Arial -Align Justify;
                $pscriboDocument.Styles['TestStyle'].Align | Should Be 'Justify';
            }

            It 'sets default style to "Normal".' {
                $pscriboDocument.DefaultStyle = $null;
                $pscriboDocument.DefaultStyle | Should Be $null;
                Style -Name Normal -Font Tahoma -Size 10 -Default;
                $pscriboDocument.DefaultStyle | Should Be 'Normal';
            }

             It 'test case insensitive style names/Ids.' {
                Style -Name normal -Font Tahoma -Size 10 -Default;
                $pscriboDocument.DefaultStyle | Should Be 'Normal';
            }

        } #end context by named parameter

        Context 'By Positional Parameter' {

            It 'sets style using positional -Id and -Font parameters.' {
                Style 'Test Style';
                $pscriboDocument.Styles['TestStyle'].Name | Should BeExactly 'Test Style';
                $pscriboDocument.Styles['TestStyle'].Id | Should BeExactly 'TestStyle';
            }

            It 'set style font size by positional -Size parameter.' {
                Style 'Test Style' 16;
                $pscriboDocument.Styles['TestStyle'].Size | Should Be 16;
            }

            It 'defaults style colour to black by color name.' {
                Style 'Test Style';
                $pscriboDocument.Styles['TestStyle'].Color | Should BeExactly '000000';
            }

            It 'defaults background colour to null.' {
                Style 'Test Style';
                $pscriboDocument.Styles['TestStyle'].BackgroundColor | Should BeNullOrEmpty;
            }

            It 'defaults text alignment to left.' {
                Style -Name 'Test Style' -Font Arial;
                $pscriboDocument.Styles['TestStyle'].Align | Should Be 'Left';
            }

            It 'tests total syle count is 2.' {
                # Should be 'Test Style' and 'Default/default'
                $pscriboDocument.Styles.Count | Should Be ($pscriboDocumentStyleCount +1);
            }

            It 'throws with invalid html color.' {
                { Style Id 'Test Style' -Color xyz } | Should Throw;
            }

            It 'throws with invalid html background color.' {
                { Style Id 'Test Style' -BackgroundColor xyz } | Should Throw;
            }

            It 'tests valid html color code' {
                Test-PscriboStyleColor '012345' | Should Be $true;
                Test-PscriboStyleColor '#012345' | Should Be $true;
                Test-PscriboStyleColor '#D3D' | Should Be $true;
                Test-PscriboStyleColor D3D | Should Be $true;
            }

            It 'tests invalid length html color code' {
                Test-PscriboStyleColor abcd | Should Be $false;
                Test-PscriboStyleColor 1abcdef | Should Be $false;
            }

            It 'tests invalid html color code' {
                Test-PscriboStyleColor ghi | Should Be $false;
            }

        } #end context by positional parameter

    } #end describe style

    Describe 'Style\Test-PScriboStyleColor' {

            It 'tests valid html color code.' {
                Test-PscriboStyleColor -Color '012345' | Should Be $true;
                Test-PscriboStyleColor -Color '#012345' | Should Be $true;
                Test-PscriboStyleColor -Color '#D3D' | Should Be $true;
                Test-PscriboStyleColor -Color D3D | Should Be $true;
            }

            It 'tests invalid length html color code.' {
                Test-PscriboStyleColor -Color abcd | Should Be $false;
                Test-PscriboStyleColor -Color 1abcdef | Should Be $false;
            }

            It 'tests invalid html color code.' {
                Test-PscriboStyleColor -Color ghi | Should Be $false;
            }
    }

    Describe 'Style\Set-PScriboStyle' {
        $pscriboDocument = Document 'ScaffoldDocument' {};
        Style -Name 'MyCustomStyle' -Size 16;

        Context 'By Row.' {
            It 'sets row style by reference.' {
                ($service = Get-Service | Select -Last 1) | Set-Style -Style 'MyCustomStyle';
                $service.__Style | Should Be 'MyCustomStyle';
            }
            It 'sets row style by pipeline.' {
                $service = Get-Service | Select -Last 1 | Set-Style -Style 'MyCustomStyle' -PassThru;
                $service.__Style | Should Be 'MyCustomStyle';
            }
        } #end context by row

        Context 'By Cell.' {
            It 'sets cell style on a single property by reference.' {
                ($service = Get-Service | Select -Last 1) | Set-Style -Style 'MyCustomStyle' -Property Status;
                $service.Status__Style | Should Be 'MyCustomStyle';
            }
            It 'sets cell style on a single property by pipeline.' {
                $service = Get-Service | Select -Last 1 | Set-Style -Style 'MyCustomStyle' -Property Status -PassThru;
                $service.Status__Style | Should Be 'MyCustomStyle';
            }
            It 'sets cell style on a multiple properties by reference.' {
                ($service = Get-Service | Select -Last 1) | Set-Style -Style 'MyCustomStyle' -Property Status,Name;
                $service.Status__Style | Should Be 'MyCustomStyle';
                $service.Name__Style | Should Be 'MyCustomStyle';
            }
            It 'sets cell style on a multiple properties by pipeline.' {
                $service = Get-Service | Select -Last 1 | Set-Style -Style 'MyCustomStyle' -Property Status,Name -PassThru;
                $service.Status__Style | Should Be 'MyCustomStyle';
                $service.Name__Style | Should Be 'MyCustomStyle';
            }

        } #end context by cell
    } #end describe set-style

} #end inmodulescope

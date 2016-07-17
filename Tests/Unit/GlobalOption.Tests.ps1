$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe "GlobalOption\GlobalOption" {
        $pscriboDocument = Document 'Test' {}

        It 'sets default space separator to "".' {
            GlobalOption;
            $pscriboDocument.Options['SpaceSeparator'] | Should Be $null;
        }

        It 'defaults to 25.4mm (1 inch) margin.' {
            GlobalOption;
            $pscriboDocument.Options['MarginTop'] | Should Be 25.4;
            $pscriboDocument.Options['MarginBottom'] | Should Be 25.4;
            $pscriboDocument.Options['MarginLeft'] | Should Be 25.4;
            $pscriboDocument.Options['MarginRight'] | Should Be 25.4;
        }

        It 'defaults to "Calibri","Candara","Segoe","Segoe UI","Optima","Arial","Sans-Serif" fonts' {
            GlobalOption;
            $pscriboDocument.Options['DefaultFont'].Count | Should Be 7;
        }

        It 'defaults to A4 page size.' {
            GlobalOption;
            $pscriboDocument.Options['PageWidth'] | Should Be 210;
            $pscriboDocument.Options['PageHeight'] | Should Be 297;
        }

        It 'sets custom space separator to "_".' {
            GlobalOption -SpaceSeparator '_';
            $pscriboDocument.Options['SpaceSeparator'] | Should Be '_';
        }

        It 'sets uppercase header flag.' {
            $pscriboDocument.Options['ForceUppercaseHeader'] | Should Be $null;
            GlobalOption -ForceUppercaseHeader;
            $pscriboDocument.Options['ForceUppercaseHeader'] | Should Be $true;
        }

        It 'sets uppercase section flag.' {
            $pscriboDocument.Options['ForceUppercaseSection'] | Should Be $null;
            GlobalOption -ForceUppercaseSection;
            $pscriboDocument.Options['ForceUppercaseSection'] | Should Be $true;
        }

        It 'sets section numbering flag.' {
            $pscriboDocument.Options['EnableSectionNumbering'] | Should Be $null;
            GlobalOption -EnableSectionNumbering;
            $pscriboDocument.Options['EnableSectionNumbering'] | Should Be $true;
        }

        It 'sets page size to US Letter.' {
            GlobalOption -PageSize Letter;
            $pscriboDocument.Options['PageWidth'] | Should Be 215.9;
            $pscriboDocument.Options['PageHeight'] | Should Be 279.4;
        }

        It 'sets page size to US Legal' {
            GlobalOption -PageSize Legal;
            $pscriboDocument.Options['PageWidth'] | Should Be 215.9;
            $pscriboDocument.Options['PageHeight'] | Should Be 355.6;
        }

        It 'sets page orientation to US Legal Portrait' {
            GlobalOption -PageSize Legal -Orientation Portrait;
            $pscriboDocument.Options['PageWidth'] | Should Be 215.9;
            $pscriboDocument.Options['PageHeight'] | Should Be 355.6;
        }

        It 'sets page orientation to US Legal Landscape' {
            GlobalOption -PageSize Legal -Orientation Landscape;
            $pscriboDocument.Options['PageHeight'] | Should Be 215.9;
            $pscriboDocument.Options['PageWidth'] | Should Be 355.6;
        }

        It 'sets page margin to 1/2 inch using 36pt.' {
            GlobalOption -Margin 36;
            $pscriboDocument.Options['MarginTop'] | Should Be 12.7;
            $pscriboDocument.Options['MarginBottom'] | Should Be 12.7;
            $pscriboDocument.Options['MarginLeft'] | Should Be 12.7;
            $pscriboDocument.Options['MarginRight'] | Should Be 12.7;
        }

    } #end Describe GlobalOption

} #end inmodulescope

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe "Export-Document\Export-Document" {

        $document = Document 'ScaffoldDocument' {};

        It "calls single output formatter." {
            Mock OutText -Verifiable { }
            $document | Export-Document -Format Text;
            Assert-VerifiableMock;
        }

        It "calls multiple output formatters." {
            Mock OutText -Verifiable { }
            Mock OutXml -Verifiable { }
            $document | Export-Document -Format Text,XML;
            Assert-VerifiableMock;
        }

        It "throws on invalid qualified directory path." {
            { $document | Export-Document -Format Text -Path TestDrive:\ThisShouldHopefullyNotExist -ErrorAction SilentlyContinue } |
                Should Throw;
        }

        It "does not throw on valid qualified path." {
            { $document | Export-Document -Format Text -Path TestDrive:\ -ErrorAction SilentlyContinue } |
                Should Not Throw;
        }

        It "throws on invalid relative directory path." {
            { $document | Export-Document -Format Text -Path .\ThisShouldHopefullyNotExist -ErrorAction SilentlyContinue } |
                Should Throw;
        }

        It "does not throw on valid relative path." {
            { $document | Export-Document -Format Text -Path . -ErrorAction SilentlyContinue } |
                Should Not Throw;
        }

        It "calls single output formatter." {
            Mock OutText -Verifiable { }
            Export-Document -Document $document -Format Text;
            Assert-VerifiableMock;
        }

        It "calls multiple output formatters." {
            Mock OutText -Verifiable { }
            Mock OutXml -Verifiable { }
            Export-Document -Document $document -Format Text,XML;
            Assert-VerifiableMock;
        }

        It "throws on invalid qualified directory path." {
            { Export-Document -Document $document -Format Text -Path TestDrive:\ThisShouldHopefullyNotExist -ErrorAction SilentlyContinue } |
                Should Throw;
        }

        It "does not throw on valid qualified path." {
            { Export-Document -Document $document -Format Text -Path TestDrive:\ -ErrorAction SilentlyContinue } |
                Should Not Throw;
        }

        It "throws on invalid relative directory path." {
            { Export-Document -Document $document -Format Text -Path .\ThisShouldHopefullyNotExist -ErrorAction SilentlyContinue } |
                Should Throw;
        }

        It "does not throw on valid relative path." {
            { Export-Document -Document $document -Format Text -Path . -ErrorAction SilentlyContinue } |
                Should Not Throw;
        }

        It "throws on invalid chars in path." {
            { Export-Document -Document $document -Format Text -Path "C:\AppData\Loc$([char]0)al\" -ErrorAction SilentlyContinue } |
                Should Throw;
        }

    }

} #end inmodulescope

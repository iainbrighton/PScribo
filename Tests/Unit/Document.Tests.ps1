$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Document' {

        It 'returns a PSCustomObject object' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.GetType().Name | Should Be 'PSCustomObject'
        }

        It 'creates a PScribo.Document type' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.Type | Should Be 'PScribo.Document'
        }

        It 'creates a Document by named -Name parameter' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.Name | Should BeExactly 'Test Document'
        }

        It 'throws when a Name of a Document contains incorrect chars' {
            { Document -Name [String]"Test-File-201606$([char]0)08-1315.txt" -ScriptBlock {} } | Should Throw
        }

        It 'default Document.Options type should be hashtable' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.Options.GetType() | Should Be 'Hashtable'
        }

        It 'default Document.Styles type should be hashtable' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.Styles.GetType() | Should Be 'Hashtable'
        }

        It 'default Document.Style should be named "Normal"' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.DefaultStyle | Should BeExactly 'Normal'
        }

        It 'default Document.TableStyle should be named "TableDefault"' {
            $document = Document -Name 'Test Document' -ScriptBlock { }
            $document.DefaultTableStyle | Should BeExactly 'TableDefault'
        }

        It 'creates a Document by positional -Name and -ScriptBlock parameters' {
            $document = Document 'Test Document' { }
            $document.Name | Should BeExactly 'Test Document'
        }

        It 'creates a concatenated document Id by positional -Name and -ScriptBLock parameters' {
            $document = Document 'Test Document' { }
            $document.Id | Should BeExactly 'TESTDOCUMENT'
        }
    }
}

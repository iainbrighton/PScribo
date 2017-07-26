$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Document\Document' {

        It 'returns a PSCustomObject object.' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.GetType().Name | Should Be 'PSCustomObject';
        }

        It 'creates a PScribo.Document type.' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.Type | Should Be 'PScribo.Document';
        }

        It 'creates a Document by named -Name parameter.' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.Name | Should BeExactly 'Test Document';
        }

        It 'throws when a Name of a Document contains incorrect chars.'{
            { Document -Name [String]"Test-File-201606$([char]0)08-1315.txt" -ScriptBlock {} } | Should Throw;
        }

        It 'default Document.Options type should be hashtable.' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.Options.GetType() | Should Be 'Hashtable';
        }

        It 'default Document.Styles type should be hashtable.' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.Styles.GetType() | Should Be 'Hashtable';
        }

        It 'default Document.Style should be named "Normal".' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.DefaultStyle | Should BeExactly 'Normal';
        }

        It 'default Document.TableStyle should be named "TableDefault".' {
            $document = Document -Name 'Test Document' -ScriptBlock { };
            $document.DefaultTableStyle | Should BeExactly 'TableDefault';
        }

        It 'creates a Document by positional -Name and -ScriptBlock parameters.' {
            $document = Document 'Test Document' { };
            $document.Name | Should BeExactly 'Test Document';
        }

        It 'creates a concatenated document Id by positional -Name and -ScriptBLock parameters.' {
            $document = Document 'Test Document' { };
            $document.Id | Should BeExactly 'TESTDOCUMENT';
        }

    } #end describe Document

} #end inmodulescope

<#
Code coverage report:
Covered 65.22% of 69 analyzed commands in 2 files.

Missed commands:

File                  Function                    Line Command
----                  --------                    ---- -------
Document.Internal.ps1 Invoke-PScriboSectionLevel   74 if ($pscriboDocument.Options['ForceUppercaseSection']) {...
Document.Internal.ps1 Invoke-PScriboSectionLevel   75 $Section.Name = $Section.Name.ToUpper()
Document.Internal.ps1 Invoke-PScriboSectionLevel   78 $Section.Number = $Number
Document.Internal.ps1 Invoke-PScriboSectionLevel   79 $Section.Level = $Number.Split('.').Count -1
Document.Internal.ps1 Invoke-PScriboSectionLevel   81 $tocEntry = [PScustomObject] @{ Id = $Section.Id; Number = $N...
Document.Internal.ps1 Invoke-PScriboSectionLevel   81 Id = $Section.Id
Document.Internal.ps1 Invoke-PScriboSectionLevel   81 Number = $Number
Document.Internal.ps1 Invoke-PScriboSectionLevel   81 Level = $Section.Level
Document.Internal.ps1 Invoke-PScriboSectionLevel   81 Name = $Section.Name
Document.Internal.ps1 Invoke-PScriboSectionLevel   82 [ref] $null = $pscriboDocument.TOC.Add($tocEntry)
Document.Internal.ps1 Invoke-PScriboSectionLevel   84 $minorNumber = 1
Document.Internal.ps1 Invoke-PScriboSectionLevel   85 $Section.Sections
Document.Internal.ps1 Invoke-PScriboSectionLevel   86 if ($s.Type -like '*.Section' -and -not $s.IsExcluded) {...
Document.Internal.ps1 Invoke-PScriboSectionLevel   87 $sectionNumber = ('{0}.{1}' -f $Number, $minorNumber).TrimSta...
Document.Internal.ps1 Invoke-PScriboSectionLevel   87 '{0}.{1}' -f $Number, $minorNumber
Document.Internal.ps1 Invoke-PScriboSectionLevel   88 Invoke-PScriboSectionLevel -Section $s -Number $sectionNumber
Document.Internal.ps1 Invoke-PScriboSectionLevel   89 $minorNumber++
Document.Internal.ps1 Invoke-PScriboSection        96 if ($s.Type -like '*.Section') {...
Document.Internal.ps1 Invoke-PScriboSection        97 if ($pscriboDocument.Options['ForceUppercaseSection']) {...
Document.Internal.ps1 Invoke-PScriboSection        98 $s.Name = $s.Name.ToUpper()
Document.Internal.ps1 Invoke-PScriboSection       100 if (-not $s.IsExcluded) {...
Document.Internal.ps1 Invoke-PScriboSection       101 Invoke-PScriboSectionLevel -Section $s -Number $majorNumber
Document.Internal.ps1 Invoke-PScriboSection       102 $majorNumber++
Document.ps1          Document                      27 [ref] $null = $pscriboDocument.Sections.Add($result)
#>

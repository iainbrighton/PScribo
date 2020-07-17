$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    ## Scaffold root XmlDocument
    $xmlDocument = New-Object -TypeName System.Xml.XmlDocument;
    [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'));
    $element = $xmlDocument.AppendChild($xmlDocument.CreateElement('testdocument'));
    [ref] $null = $element.SetAttribute('name', 'testdocument');

    ## Scaffold document
    $pscriboDocument = Document -Name 'TestDocument' -ScriptBlock {
        Paragraph -Name 'TestParagraph';
        Get-Process | Select-Object -First 3 | Table 'TestTable';
        Section -Name 'TestSection' -ScriptBlock {
            Section -Name 'TestSubSection' { }
        }
    }

    Describe 'Plugins\Xml\Out-XmlTable' {

        $processes = Get-Process | Select -First 3;
        $tableName = 'Test Table';
        $tableColumns = @('ProcessName','SI','Id');
        $tableHeaders = @('ProcessName','SI','Id');

        Context 'By Name Parameter.' {

            It 'outputs a root XmlElement' {
                $table = $processes | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    Out-XmlTable

                    $table.LocalName | Should BeExactly ($tableName.Replace(' ','').ToLower())
                $table.Name | Should BeExactly $tableName

            }

            It 'outputs a XmlElement for each table row' {
                $table = $processes | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    Out-XmlTable

                $table.ChildNodes.Count | Should Be $processes.Count
            }

            It 'outputs a XmlElement for each table column' {
                $table = $processes | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    Out-XmlTable

                $table.FirstChild.ChildNodes.Count | Should Be $tableColumns.Count
            }

            It 'creates a name attribute when headers contain spaces' {
                $table = $processes | Table -Name $tableName -Columns $tableColumns -Headers $tableHeaders |
                    Out-XmlTable

                for ($i = 0; $i -lt $table.FirstChild.ChildNodes.Count; $i++)
                {
                    $table.FirstChild.ChildNodes[$i].LocalName | Should BeExactly $tableColumns[$i].ToLower()
                    if ($tableHeaders[$i].Contains(' '))
                    {
                        $table.FirstChild.ChildNodes[$i].Name | Should BeExactly $tableHeaders[$i]
                    }
                }
            }
        }

        Context 'By Postional Parameter.' {

            It 'outputs a root XmlElement' {
                $table = $processes | Table $tableName $tableColumns $tableHeaders |
                    Out-XmlTable

                $table.LocalName | Should BeExactly ($tableName.Replace(' ','').ToLower())
                $table.Name | Should BeExactly $tableName
                $table.ChildNodes.Count | Should Be 3
            }

            It 'outputs a XmlElement for each table row' {
                $table = $processes | Table $tableName $tableColumns $tableHeaders |
                    Out-XmlTable

                $table.ChildNodes.Count | Should Be $processes.Count
            }

            It 'outputs a XmlElement for each table column' {
                $table = $processes | Table $tableName $tableColumns $tableHeaders |
                    Out-XmlTable

                $table.FirstChild.ChildNodes.Count | Should Be $tableColumns.Count
            }

            It 'creates a name attribute when headers contain spaces' {
                $table = $processes | Table $tableName $tableColumns $tableHeaders |
                    Out-XmlTable

                for ($i = 0; $i -lt $table.FirstChild.ChildNodes.Count; $i++)
                {
                    $table.FirstChild.ChildNodes[$i].LocalName | Should BeExactly $tableColumns[$i].ToLower()
                    if ($tableHeaders[$i].Contains(' '))
                    {
                        $table.FirstChild.ChildNodes[$i].Name | Should BeExactly $tableHeaders[$i]
                    }
                }
            }
        }
    }
}

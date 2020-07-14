$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Out-HtmlTable' {

        $script:currentPageNumber = 1

        Context 'Table' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { }
                $Document = $pscriboDocument
                $processes = Get-Process | Select-Object -First 3
                $table = $processes | Table -Name 'Test Table' | Out-HtmlTable
                [Xml] $html = $table.Replace('&','&amp;')
            }

            It 'creates default table class of tabledefault' {
                $html.Div.Table.Class | Should BeExactly 'tabledefault'
            }

            It 'creates table headings row' {
                $html.Div.Table.Thead | Should Not BeNullOrEmpty
            }

            It 'creates column for each object property' {
                $html.Div.Table.Thead.Tr.Th.Count | Should Be ($processes | Get-Member -MemberType Properties).Count
            }

            It 'creates a row for each object' {
                $html.Div.Table.Tbody.Tr.Count | Should Be $processes.Count
            }
        }

        Context 'List' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { }
                $Document = $pscriboDocument

                $processes = Get-Process | Select-Object -First 1
                $table = $processes | Table -Name 'Test Table' -List | Out-HtmlTable

                [Xml] $html = $table.Replace('&','&amp;').Replace('<p />','')
            }

            It 'creates no table heading row' {
                ## Fix Set-StrictMode
                $html.Div.Table.PSObject.Properties['Thead'] | Should BeNullOrEmpty
            }

            It 'creates default table class of tabledefault' {
                $html.Div.Table.Class | Should BeExactly 'tabledefault'
            }

            It 'creates a two column table' {
                $html.Div.Table.Tbody.Tr[0].ChildNodes.Count | Should Be 2
            }

            It 'creates a row for each object property' {
                $html.Div.Table.Tbody.Tr.Count | Should Be ($processes | Get-Member -MemberType Properties).Count
            }

        } #end context List

        Context 'New Lines' {

            BeforeEach {
                ## Scaffold new document to initialise options/styles
                $pscriboDocument = Document -Name 'Test' -ScriptBlock { }
                $Document = $pscriboDocument
            }

            It 'creates a tabular table cell with an embedded new line' {

                $licenses = 'Standard{0}Professional{0}Enterprise' -f [System.Environment]::NewLine
                $expected = '<td>Standard<br />Professional<br />Enterprise</td>'
                $newLineTable = [PSCustomObject] @{ 'Licenses' = $licenses; }

                $table = $newLineTable | Table -Name 'Test Table' | Out-HtmlTable

                [Xml] $html = $table.Replace('&','&amp;')
                $html.OuterXml | Should Match $expected
            }

            It 'creates a list table cell with an embedded new line' {

                $licenses = 'Standard{0}Professional{0}Enterprise' -f [System.Environment]::NewLine
                $expected = '<td>Standard<br />Professional<br />Enterprise</td>'
                $newLineTable = [PSCustomObject] @{ 'Licenses' = $licenses; }

                $table = $newLineTable | Table -Name 'Test Table' | Out-HtmlTable

                [Xml] $html = $table.Replace('&','&amp;')
                $html.OuterXml | Should Match $expected
            }
        }
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Table' {

        $pscriboDocument = Document 'ScaffoldDocument' {}

        Context 'InputObject, By Named Parameter' {

            $processes = Get-Process | Select-Object -First 3

            It 'returns a PSCustomObject object' {
                $table = $processes | Table

                $table.GetType().Name | Should Be 'PSCustomObject'
            }

            It 'creates a PScribo.Table type' {
                $table = $processes | Table

                $table.Type | Should Be 'PScribo.Table'
            }

            It 'creates a table without a name parameter' {
                $table = $processes | Table

                $table.Id | Should Not Be $null
                $table.Name | Should Not Be $null
            }

            It 'defaults to table style "TableDefault"' {
                $table = $processes | Table

                $table.Style | Should BeExactly 'TableDefault'
            }

            It 'creates a table with -List parameter' {
                $table = $processes | Table -List

                $table.IsList | Should Be $true
            }

            It 'creates a table by named -Name parameter' {
                $table = $processes | Table -Name 'TestTable'

                $table.Id | Should BeExactly 'TESTTABLE'
                $table.Name | Should BeExactly 'TestTable'
            }

            It 'creates a table by named -Name parameter with a space' {
                $table = $processes | Table -Name 'Test Table'

                $table.Id | Should BeExactly 'TESTTABLE'
                $table.Name | Should BeExactly 'Test Table'
            }

            It 'creates a table by named -Columns parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )

                $table = $processes | Table -Columns $columns

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by named -Columns and -Headers parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'Name', 'Session Id', 'Process Id' )

                $table = $processes | Table -Name 'TestTable' -Columns $columns -Headers $headers

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by mismatched named -Headers and -Columns parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'Name', 'Session Id', 'Process Id' )

                $table = Table -InputObject $processes -Name 'TestTable' -Headers $headers -Columns $columns -WarningAction SilentlyContinue

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by named -Columns, -Headers and -Style parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'Name', 'Session Id', 'Process Id' )
                $styleName = 'TestStyle'

                $table = $processes | Table -Name 'TestTable' -Columns $columns -Headers $headers -Style $styleName

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
                $table.Style | Should Be $styleName
            }

            It 'warns with more than 2 columns with named -List parameter' {
                $columnWidths = @(100,200,300)

                $table = $processes | Table -List -ColumnWidths $columnWidths -WarningAction SilentlyContinue

                $table.ColumnWidths | Should BeNullOrEmpty
            }

            It 'warns with mismatching columns and column widths' {
                $columnWidths = @(100,200,300,400)

                $table = $processes | Table -ColumnWidths $columnWidths -WarningAction SilentlyContinue

                $table.ColumnWidths | Should BeNullOrEmpty
            }

            It 'creates a table with specified column widths' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'Name', 'Session Id', 'Process Id' )
                $columnWidths = @(25,35,40)

                $table = $processes | Table -Name 'TestTable' -Columns $columns -Headers $headers -ColumnWidths $columnWidths

                $table.ColumnWidths[0] | Should Be $columnWidths[0]
                $table.ColumnWidths[1] | Should Be $columnWidths[1]
                $table.ColumnWidths[2] | Should Be $columnWidths[2]
            }

            It 'creates a table with embedded row style' {
                $process = $processes[1]
                Add-Member -InputObject $process -MemberType NoteProperty -Name __Style -Value 'MyRowStyle'

                $table = $processes | Table -Name 'TestTable'

                $table.Rows[1].__Style | Should Be 'MyRowStyle'
                ($table.Columns -notlike '*__Style') -as [System.Boolean] | Should Be $true
            }

            It 'creates a table with embedded cell style' {
                $process = $processes[1]
                Add-Member -InputObject $process -MemberType NoteProperty -Name Name__Style -Value 'MyCellStyle'

                $table = $processes | Table -Name 'TestTable'

                $table.Rows[1].Name__Style | Should Be 'MyCellStyle'
                ($table.Columns -notlike '*__Style') -as [System.Boolean] | Should Be $true
            }
        }

        Context 'Hashtable, By Named Parameter' {

            [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'returns a PSCustomObject object' {
                $table = Table -Hashtable $services

                $table.GetType().Name | Should Be 'PSCustomObject'
            }

            It 'creates a PScribo.Table type' {
                $table = Table -Hashtable $services

                $table.Type | Should Be 'PScribo.Table'
            }

            It 'creates a table without a name parameter' {
                $table = Table -Hashtable $services

                $table.Id | Should Not Be $null
                $table.Name | Should Not Be $null
            }

            It 'defaults to table style "TableDefault"' {
                $table = Table -Hashtable $services

                $table.Style | Should BeExactly 'TableDefault'
            }

            It 'creates a table with -List parameter' {
                $table = Table -Hashtable $services -List

                $table.IsList | Should Be $true
            }

            It 'creates a table by named -Name parameter' {
                $table = Table -Hashtable $services -Name 'TestTable'

                $table.Id | Should BeExactly 'TESTTABLE'
                $table.Name | Should BeExactly 'TestTable'
            }

            It 'creates a table by named -Name parameter with a space' {
                $table = Table -Hashtable $services -Name 'Test Table'

                $table.Id | Should BeExactly 'TESTTABLE'
                $table.Name | Should BeExactly 'Test Table'
            }

            It 'creates a table by named -Columns parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )

                $table = Table -Hashtable $services -Columns $columns

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by named -Columns and -Headers parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name', 'Display Name' )

                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by named -Columns, -Headers and -Style parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name', 'Display Name' )
                $styleName = 'TestStyle'

                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers -Style $styleName

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
                $table.Style | Should Be $styleName
            }

            It 'creates a table by mismatched named -Headers and -Columns parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name' )

                $table = Table -Hashtable $services -Name 'TestTable' -Headers $headers -Columns $columns -WarningAction SilentlyContinue

                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'warns with more than 2 columns with named -List parameter' {
                $columnWidths = @(100,200,300)

                $table = Table -Hashtable $services -List -ColumnWidths $columnWidths -WarningAction SilentlyContinue

                $table.ColumnWidths | Should BeNullOrEmpty
            }

            It 'creates a table with specified column widths' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; ServiceName = 'Test 1'; DisplayName = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; ServiceName = 'Test 3'; DisplayName = 'Test Service 3'; __Style = 'MyRowStyle'; }
                    [ordered] @{ Name = 'TestService2'; ServiceName = 'Test 2'; DisplayName = 'Test Service 2'; }
                )
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name', 'Display Name' )
                $columnWidths = @(25,35,40)

                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers -ColumnWidths $columnWidths

                $table.ColumnWidths[0] | Should Be $columnWidths[0]
                $table.ColumnWidths[1] | Should Be $columnWidths[1]
                $table.ColumnWidths[2] | Should Be $columnWidths[2]
            }

            It 'creates a table with embedded row style' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; __Style = 'MyRowStyle'; }
                    [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
                )

                $table = Table -Hashtable $services -Name 'TestTable'

                $table.Rows[1].__Style | Should Be 'MyRowStyle'
                ($table.Columns -notlike '*__Style') -as [System.Boolean]  | Should Be $true
            }

            It 'creates a table with embedded cell style' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; Name__Style = 'MyCellStyle'; }
                    [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
                )

                $table = Table -Hashtable $services -Name 'TestTable'

                $table.Rows[1].Name__Style | Should Be 'MyCellStyle'
                ($table.Columns -notlike '*__Style') -as [System.Boolean]  | Should Be $true
            }
        }

        Context 'InputObject, By Positional Parameters' {

            $processes = Get-Process | Select-Object -First 3
            $tableName = 'TestName'

            It 'creates a table by positional -Name parameter' {
                $table = $processes | Table $tableName

                $table.Id | Should BeExactly $tableName.ToUpper()
                $table.Name | Should BeExactly $tableName
            }

            It 'creates a table by positional -Name parameter with a space' {
                $table = $processes | Table 'Test Name'

                $table.Id | Should BeExactly $tableName.ToUpper()
                $table.Name | Should BeExactly 'Test Name'
            }

            It 'creates a table by positional -Name and -Columns parameters' {
                $tableName = 'TestName'
                $columns = @( 'ProcessName', 'SI', 'Id' )

                $table = $processes | Table $tableName $columns

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by positional -Name, -Columns and -Headers parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'ProcessName', 'Session Id', 'Process Id' )

                $table = $processes | Table $tableName $columns $headers

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by positional -Name, -Columns, -Headers and -Style parameters' {
                $columns = @( 'ProcessName', 'SI', 'Id' )
                $headers = @( 'ProcessName', 'Session Id', 'Process Id' )
                $styleName = 'TestStyle'

                $table = $processes | Table $tableName $columns $headers $styleName

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
                $table.Style | Should Be $styleName
            }
        }

        Context 'Hashtable, By Positional Parameter' {

            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )
            $tableName = 'TestName'

            It 'creates a table by positional -Name parameter' {
                $table = $services | Table $tableName

                $table.Id | Should BeExactly $tableName.ToUpper()
                $table.Name | Should BeExactly $tableName
            }

            It 'creates a table by positional -Name parameter with a space' {
                $table = $services | Table 'Test Name'

                $table.Id | Should BeExactly $tableName.ToUpper()
                $table.Name | Should BeExactly 'Test Name'
            }

            It 'creates a table by positional -Name and -Columns parameters' {
                $tableName = 'TestName'
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )

                $table = $services | Table $tableName $columns

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by positional -Name, -Columns and -Headers parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name', 'Display Name' )

                $table = $services | Table $tableName $columns $headers

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
            }

            It 'creates a table by positional -Name, -Columns, -Headers and -Style parameters' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' )
                $headers = @( 'Name', 'Service Name', 'Display Name' )
                $styleName = 'TestStyle'

                $table = $services | Table $tableName $columns $headers $styleName

                $table.Name | Should Be $tableName
                $table.Columns.Count | Should Be 3
                $table.Rows.Count | Should Be 3
                $table.Style | Should Be $styleName
            }
        }
    }
}

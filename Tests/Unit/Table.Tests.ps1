$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Table\New-PScriboTableRow' {

        Context 'By System.Object.' {

            $service = Get-Service | Select -First 1;

            It 'returns a PSCustomObject.' {
                $row = New-PScriboTableRow -InputObject $service;
                $row.GetType().Name | Should Be 'PSCustomObject';
                $row.PSObject.Properties['__Style'].Value | Should BeNullOrEmpty;
            }

            It 'creates a row by a single named -Properties parameter.' {
                $row = $service | New-PScriboTableRow -Properties 'DisplayName';
                ($row | Get-Member -MemberType Properties).Count | Should Be 2; ## Row __Style property added
            }

            It 'creates a row by named -Properties.' {
                $serviceProperties = $service | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name;
                $row = New-PScriboTableRow -InputObject $service -Properties $serviceProperties;
                ($row | Get-Member -MemberType Properties).Count | Should Be ($serviceProperties.Count +1); ## Row __Style property added
            }

            It 'creates a row by named -Properties and -Headers parameters.' {
                $properties = 'Name', 'ServiceName', 'DisplayName';
                $headers = 'Name', 'Service Name', 'Display Name';
                $row = New-PScriboTableRow -InputObject $service -Properties $properties -Headers $headers;
                ($row | Get-Member -MemberType Properties).Count | Should Be 4; ## Row __Style property added
                $row.Name | Should Not Be $null;
                $row.'Service Name' | Should Not Be $null;
                $row.'Display Name' | Should Not Be $null;
                $row.PSObject.Properties['ServiceName'] | Should BeNullOrEmpty;
                $row.PSObject.Properties['DisplayName'] | Should BeNullOrEmpty;
            }

        } #end context by System.Object

        Context 'By System.Management.Automation.PSCustomObject.' {
            $serviceName = 'TestService';
            $serviceServiceName = 'Test';
            $serviceDisplayName = 'Test Service';
            $service = [PSCustomObject] @{
                Name = $serviceName;
                ServiceName = $serviceServiceName;
                DisplayName = $serviceDisplayName;
            }

            It 'returns a PSCustomObject.' {
                $row = New-PScriboTableRow -InputObject $service;
                $row.GetType().Name | Should Be 'PSCustomObject';
            }

            It 'creates a table row by a single named -Properties parameter.' {
                $row = $service | New-PScriboTableRow -Properties 'DisplayName';
                ($row | Get-Member -MemberType Properties).Count | Should Be 2; ## Row __Style property added
            }

            It 'creates a table row by named -Properties.' {
                $properties = 'Name','DisplayName';
                $row = New-PScriboTableRow -InputObject $service -Properties $properties;
                ($row | Get-Member -MemberType Properties).Count | Should Be ($properties.Count +1); ## Row __Style property added
            }

             It 'creates a table row by named -Properties and -Headers parameters.' {
                $properties = 'Name', 'ServiceName', 'DisplayName';
                $headers = 'Name', 'Service Name', 'Display Name';
                $row = New-PScriboTableRow -InputObject $service -Properties $properties -Headers $headers;
                ($row | Get-Member -MemberType Properties).Count | Should Be 4; ## Row __Style property added
                $row.Name | Should Not Be $null;
                $row.'Service Name' | Should Not Be $null;
                $row.'Display Name' | Should Not Be $null;
                $row.PSObject.Properties['ServiceName'] | Should BeNullOrEmpty;
                $row.PSObject.Properties['DisplayName'] | Should BeNullOrEmpty;
            }

        } #end context By System.Management.Automation.PSCustomObject

        Context 'By Hashtable.' {
            $serviceName = 'TestService';
            $serviceServiceName = 'Test';
            $serviceDisplayName = 'Test Service';
            $service = [Ordered] @{ Name = $serviceName; ServiceName = $serviceServiceName; DisplayName = $serviceDisplayName; }

            It 'returns a PSCustomObject object.' {
                $row = New-PScriboTableRow -Hashtable $service;
                $row.GetType().Name | Should Be 'PSCustomObject';
                $row.__Style | Should BeNullOrEmpty;
            }

            It 'creates a table row without spaces.' {
                $row = New-PScriboTableRow -Hashtable $service;
                ($row | Get-Member -MemberType Properties).Count | Should Be 4;
                $row.Name | Should Be $serviceName;
                $row.ServiceName | Should Be $serviceServiceName;
                $row.DisplayName | Should Be $serviceDisplayName;
            }

            It 'creates a table row with spaces.' {
                $service = [Ordered] @{ Name = $serviceName; 'Service Name' = $serviceServiceName; 'Display Name' = $serviceDisplayName; }
                $row = New-PScriboTableRow -Hashtable $service;
                ($row | Get-Member -MemberType Properties).Count | Should Be 4; ## Row __Style property added
                $row.'Name' | Should Be $serviceName;
                $row.'Service Name' | Should Be $serviceServiceName;
                $row.'Display Name' | Should Be $serviceDisplayName;
            }

        } #end context by hashtable

    } #end describe

    Describe 'Table\Table' {
        $pscriboDocument = Document 'ScaffoldDocument' {};

        Context 'InputObject, By Named Parameter.' {
            $services = Get-Service | Select -First 3;

            It 'returns a PSCustomObject object.' {
                $table = $services | Table;
                $table.GetType().Name | Should Be 'PSCustomObject';
            }

            It 'creates a PScribo.Table type.' {
                $table = $services | Table;
                $table.Type | Should Be 'PScribo.Table';
            }

            It 'creates a table without a name parameter.' {
                $table = $services | Table;
                $table.Id | Should Not Be $null;
                $table.Name | Should Not Be $null;
            }

            It 'defaults to table style "TableDefault".' {
                $table = $services | Table;
                $table.Style | Should BeExactly 'TableDefault';
            }

            It 'creates a table with -List parameter.' {
                $table = $services | Table -List;
                $table.List | Should Be $true;
            }

            It 'creates a table by named -Name parameter.' {
                $table = $services | Table -Name 'TestTable';
                $table.Id | Should BeExactly 'TESTTABLE';
                $table.Name | Should BeExactly 'TestTable';
            }

            It 'creates a table by named -Name parameter with a space.' {
                $table = $services | Table -Name 'Test Table';
                $table.Id | Should BeExactly 'TESTTABLE';
                $table.Name | Should BeExactly 'Test Table';
            }

            It 'creates a table by named -Columns parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $table = $services | Table -Columns $columns;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by named -Columns and -Headers parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $table = $services | Table -Name 'TestTable' -Columns $columns -Headers $headers;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by mismatched named -Headers and -Columns parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name' );
                $table = Table -InputObject $services -Name 'TestTable' -Headers $headers -Columns $columns -WarningAction SilentlyContinue;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by named -Columns, -Headers and -Style parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $styleName = 'TestStyle';
                $table = $services | Table -Name 'TestTable' -Columns $columns -Headers $headers -Style $styleName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
                $table.Style | Should Be $styleName;
            }

            It 'warns with more than 2 columns with named -List parameter.' {
                $columnWidths = @(100,200,300);
                $table = $services | Table -List -ColumnWidths $columnWidths -WarningAction SilentlyContinue;
                $table.ColumnWidths | Should BeNullOrEmpty;
            }

            It 'warns with mismatching columns and column widths.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $columnWidths = @(100,200,300,400);
                $table = $services | Table -ColumnWidths $columnWidths -WarningAction SilentlyContinue;
                $table.ColumnWidths | Should BeNullOrEmpty;
            }

            It 'creates a table with specified column widths.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $columnWidths = @(25,35,40);
                $table = $services | Table -Name 'TestTable' -Columns $columns -Headers $headers -ColumnWidths $columnWidths;
                $table.ColumnWidths[0] | Should Be $columnWidths[0];
                $table.ColumnWidths[1] | Should Be $columnWidths[1];
                $table.ColumnWidths[2] | Should Be $columnWidths[2];
            }

            It 'creates a table with embedded row style' {
                $service = $services[1];
                Add-Member -InputObject $service -MemberType NoteProperty -Name __Style -Value 'MyRowStyle';
                $table = $services | Table -Name 'TestTable';
                $table.Rows[1].__Style | Should Be 'MyRowStyle';
                ($table.Columns -notlike '*__Style') -as [System.Boolean] | Should Be $true;
            }

            It 'creates a table with embedded cell style' {
                $service = $services[1];
                Add-Member -InputObject $service -MemberType NoteProperty -Name Name__Style -Value 'MyCellStyle';
                $table = $services | Table -Name 'TestTable';
                $table.Rows[1].Name__Style | Should Be 'MyCellStyle';
                ($table.Columns -notlike '*__Style') -as [System.Boolean] | Should Be $true;
            }

        } #end context InputObject, By Named Parameter

        Context 'Hashtable, By Named Parameter.' {
            [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )

            It 'returns a PSCustomObject object.' {
                $table = Table -Hashtable $services;
                $table.GetType().Name | Should Be 'PSCustomObject';
            }

            It 'creates a PScribo.Table type.' {
                $table = Table -Hashtable $services;
                $table.Type | Should Be 'PScribo.Table';
            }

            It 'creates a table without a name parameter.' {
                $table = Table -Hashtable $services;
                $table.Id | Should Not Be $null;
                $table.Name | Should Not Be $null;
            }

            It 'defaults to table style "TableDefault".' {
                $table = Table -Hashtable $services;
                $table.Style | Should BeExactly 'TableDefault';
            }

            It 'creates a table with -List parameter.' {
                $table = Table -Hashtable $services -List;
                $table.List | Should Be $true;
            }

            It 'creates a table by named -Name parameter.' {
                $table = Table -Hashtable $services -Name 'TestTable';
                $table.Id | Should BeExactly 'TESTTABLE';
                $table.Name | Should BeExactly 'TestTable';
            }

            It 'creates a table by named -Name parameter with a space.' {
                $table = Table -Hashtable $services -Name 'Test Table';
                $table.Id | Should BeExactly 'TESTTABLE';
                $table.Name | Should BeExactly 'Test Table';
            }

            It 'creates a table by named -Columns parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $table = Table -Hashtable $services -Columns $columns;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by named -Columns and -Headers parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by named -Columns, -Headers and -Style parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $styleName = 'TestStyle';
                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers -Style $styleName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
                $table.Style | Should Be $styleName;
            }

            It 'creates a table by mismatched named -Headers and -Columns parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name' );
                $table = Table -Hashtable $services -Name 'TestTable' -Headers $headers -Columns $columns -WarningAction SilentlyContinue;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'warns with more than 2 columns with named -List parameter.' {
                $columnWidths = @(100,200,300);
                $table = Table -Hashtable $services -List -ColumnWidths $columnWidths -WarningAction SilentlyContinue;
                $table.ColumnWidths | Should BeNullOrEmpty;
            }

            It 'creates a table with specified column widths.' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; ServiceName = 'Test 1'; DisplayName = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; ServiceName = 'Test 3'; DisplayName = 'Test Service 3'; __Style = 'MyRowStyle'; }
                    [ordered] @{ Name = 'TestService2'; ServiceName = 'Test 2'; DisplayName = 'Test Service 2'; }
                )
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $columnWidths = @(25,35,40);
                $table = Table -Hashtable $services -Name 'TestTable' -Columns $columns -Headers $headers -ColumnWidths $columnWidths;
                $table.ColumnWidths[0] | Should Be $columnWidths[0];
                $table.ColumnWidths[1] | Should Be $columnWidths[1];
                $table.ColumnWidths[2] | Should Be $columnWidths[2];
            }

            It 'creates a table with embedded row style' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; __Style = 'MyRowStyle'; }
                    [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
                )
                $table = Table -Hashtable $services -Name 'TestTable';
                $table.Rows[1].__Style | Should Be 'MyRowStyle';
                ($table.Columns -notlike '*__Style') -as [System.Boolean]  | Should Be $true;
            }

            It 'creates a table with embedded cell style' {
                [System.Collections.Specialized.OrderedDictionary[]] $services = @(
                    [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                    [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; Name__Style = 'MyCellStyle'; }
                    [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
                )
                $table = Table -Hashtable $services -Name 'TestTable';
                $table.Rows[1].Name__Style | Should Be 'MyCellStyle';
                ($table.Columns -notlike '*__Style') -as [System.Boolean]  | Should Be $true;
            }

        } #end context Hashtable, By Named Parameter

        Context 'InputObject, By Positional Parameters.' {
            $services = Get-Service | Select -First 3;
            $tableName = 'TestName';

            It 'creates a table by positional -Name parameter.' {
                $table = $services | Table $tableName;
                $table.Id | Should BeExactly $tableName.ToUpper();
                $table.Name | Should BeExactly $tableName;
            }

            It 'creates a table by positional -Name parameter with a space.' {
                $table = $services | Table 'Test Name';
                $table.Id | Should BeExactly $tableName.ToUpper();
                $table.Name | Should BeExactly 'Test Name';
            }

            It 'creates a table by positional -Name and -Columns parameters.' {
                $tableName = 'TestName';
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $table = $services | Table $tableName $columns;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by positional -Name, -Columns and -Headers parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $table = $services | Table $tableName $columns $headers;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by positional -Name, -Columns, -Headers and -Style parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $styleName = 'TestStyle';
                $table = $services | Table $tableName $columns $headers $styleName;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
                $table.Style | Should Be $styleName;
            }

        } #end context InputObject, By Positional Parameter

        Context 'Hashtable, By Positional Parameter.' {
            $services = @(
                [ordered] @{ Name = 'TestService1'; 'Service Name' = 'Test 1'; 'Display Name' = 'Test Service 1'; }
                [ordered] @{ Name = 'TestService3'; 'Service Name' = 'Test 3'; 'Display Name' = 'Test Service 3'; }
                [ordered] @{ Name = 'TestService2'; 'Service Name' = 'Test 2'; 'Display Name' = 'Test Service 2'; }
            )
            $tableName = 'TestName';

            It 'creates a table by positional -Name parameter.' {
                $table = $services | Table $tableName;
                $table.Id | Should BeExactly $tableName.ToUpper();
                $table.Name | Should BeExactly $tableName;
            }

            It 'creates a table by positional -Name parameter with a space.' {
                $table = $services | Table 'Test Name';
                $table.Id | Should BeExactly $tableName.ToUpper();
                $table.Name | Should BeExactly 'Test Name';
            }

            It 'creates a table by positional -Name and -Columns parameters.' {
                $tableName = 'TestName';
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $table = $services | Table $tableName $columns;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by positional -Name, -Columns and -Headers parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $table = $services | Table $tableName $columns $headers;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
            }

            It 'creates a table by positional -Name, -Columns, -Headers and -Style parameters.' {
                $columns = @( 'Name', 'ServiceName', 'DisplayName' );
                $headers = @( 'Name', 'Service Name', 'Display Name' );
                $styleName = 'TestStyle';
                $table = $services | Table $tableName $columns $headers $styleName;
                $table.Name | Should Be $tableName;
                $table.Columns.Count | Should Be 3;
                $table.Rows.Count | Should Be 3;
                $table.Style | Should Be $styleName;
            }

        } #end context Hashtable, By Positional Parameter

    } #end describe table

} #end inmodulescope
<#
Missed commands:

File      Function Line Command
----      -------- ---- -------
table.ps1 Table      43 Write-Warning ('Table headers specified with no table properties. Headers will be ignored.')
table.ps1 Table      43 'Table headers specified with no table properties. Headers will be ignored.'
table.ps1 Table      44 $Headers = $Columns
table.ps1 Table      65 $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Obje...
table.ps1 Table      65 $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Obje...
table.ps1 Table      65 $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Obje...
table.ps1 Table      85 Write-Warning ('Table columns and column widths do not match. Column widths will be ignored.')
table.ps1 Table      85 'Table columns and column widths do not match. Column widths will be ignored.'
#>

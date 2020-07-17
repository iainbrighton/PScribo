$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'New-PScriboTableRow' {

        Context 'By [System.Object]' {

            $process = Get-Process | Select-Object -First 1

            It 'returns a PSCustomObject' {
                $row = New-PScriboTableRow -InputObject $process

                $row.GetType().Name | Should Be 'PSCustomObject'
                $row.PSObject.Properties['__Style'].Value | Should BeNullOrEmpty
            }

            It 'creates a row by a single named -Properties parameter' {
                $row = $process | New-PScriboTableRow -Properties 'ProcessName'

                ($row | Get-Member -MemberType Properties).Count | Should Be 2 ## Row __Style property added
            }

            It 'creates a row by named -Properties' {
                $processProperties = $process |
                    Get-Member -MemberType Properties |
                        Select-Object -ExpandProperty Name

                $row = New-PScriboTableRow -InputObject $process -Properties $ProcessProperties

                ($row | Get-Member -MemberType Properties).Count | Should Be ($processProperties.Count +1) ## Row __Style property added
            }

            It 'creates a row by named -Properties and -Headers parameters' {
                $properties = 'ProcessName', 'SI', 'Id'
                $headers = 'Name', 'Session Id', 'Process Id'

                $row = New-PScriboTableRow -InputObject $process -Properties $properties -Headers $headers

                ($row | Get-Member -MemberType Properties).Count | Should Be 4 ## Row __Style property added
                $row.Name | Should Not Be $null
                $row.'Session Id' | Should Not Be $null
                $row.'Process Id' | Should Not Be $null
                $row.PSObject.Properties['SI'] | Should BeNullOrEmpty
                $row.PSObject.Properties['Id'] | Should BeNullOrEmpty
            }
        }

        Context 'By [System.Management.Automation.PSObject]' {
            $serviceName = 'TestService'
            $serviceServiceName = 'Test'
            $serviceDisplayName = 'Test Service'
            $service = [PSCustomObject] @{
                Name = $serviceName
                ServiceName = $serviceServiceName
                DisplayName = $serviceDisplayName
            }

            It 'returns a PSCustomObject' {
                $row = New-PScriboTableRow -InputObject $service

                $row.GetType().Name | Should Be 'PSCustomObject'
            }

            It 'creates a table row by a single named -Properties parameter' {
                $row = $service | New-PScriboTableRow -Properties 'DisplayName'

                ($row | Get-Member -MemberType Properties).Count | Should Be 2 ## Row __Style property added
            }

            It 'creates a table row by named -Properties' {
                $properties = 'Name','DisplayName'

                $row = New-PScriboTableRow -InputObject $service -Properties $properties

                ($row | Get-Member -MemberType Properties).Count | Should Be ($properties.Count +1) ## Row __Style property added
            }

             It 'creates a table row by named -Properties and -Headers parameters' {
                $properties = 'Name', 'ServiceName', 'DisplayName'
                $headers = 'Name', 'Service Name', 'Display Name'

                $row = New-PScriboTableRow -InputObject $service -Properties $properties -Headers $headers

                ($row | Get-Member -MemberType Properties).Count | Should Be 4 ## Row __Style property added
                $row.Name | Should Not Be $null
                $row.'Service Name' | Should Not Be $null
                $row.'Display Name' | Should Not Be $null
                $row.PSObject.Properties['ServiceName'] | Should BeNullOrEmpty
                $row.PSObject.Properties['DisplayName'] | Should BeNullOrEmpty
            }
        }

        Context 'By [System.Collections.Specialized.OrderedDictionary]' {
            $serviceName = 'TestService'
            $serviceServiceName = 'Test'
            $serviceDisplayName = 'Test Service'
            $service = [Ordered] @{ Name = $serviceName; ServiceName = $serviceServiceName; DisplayName = $serviceDisplayName; }

            It 'returns a PSCustomObject object' {
                $row = New-PScriboTableRow -Hashtable $service

                $row.GetType().Name | Should Be 'PSCustomObject'
                $row.__Style | Should BeNullOrEmpty
            }

            It 'creates a table row without spaces' {
                $row = New-PScriboTableRow -Hashtable $service

                ($row | Get-Member -MemberType Properties).Count | Should Be 4
                $row.Name | Should Be $serviceName
                $row.ServiceName | Should Be $serviceServiceName
                $row.DisplayName | Should Be $serviceDisplayName
            }

            It 'creates a table row with spaces' {
                $service = [Ordered] @{ Name = $serviceName; 'Service Name' = $serviceServiceName; 'Display Name' = $serviceDisplayName; }
                $row = New-PScriboTableRow -Hashtable $service

                ($row | Get-Member -MemberType Properties).Count | Should Be 4 ## Row __Style property added
                $row.'Name' | Should Be $serviceName
                $row.'Service Name' | Should Be $serviceServiceName
                $row.'Display Name' | Should Be $serviceDisplayName
            }
        }
    }
}

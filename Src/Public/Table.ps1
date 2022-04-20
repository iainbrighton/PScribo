function Table
{
<#
    .SYNOPSIS
        Defines a new PScribo document table.

    .PARAMETER Name

    .PARAMETER InputObject

    .PARAMETER Hashtable

    .PARAMETER Columns

    .PARAMETER ColumnWidths

    .PARAMETER Headers

    .PARAMETER Style

    .PARAMETER Width

    .PARAMETER Tabs

    .PARAMETER List

    .PARAMETER Key

    .PARAMETER Caption

    .EXAMPLE
        Table -Name 'Table 1' -InputObject $(Get-Service) -Columns 'Name','DisplayName','Status' -ColumnWidths 40,20,40
#>
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## Table name/Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Name = ([System.Guid]::NewGuid().ToString()),

        # Array of objects
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObjectList')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObjectListKey')]
        [Alias('CustomObject','Object')]
        [ValidateNotNullOrEmpty()]
        [System.Object[]] $InputObject,

        # Array of Hashtables
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Hashtable')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'HashtableList')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'HashtableListKey')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Specialized.OrderedDictionary[]] $Hashtable,

        # Array of Hashtable key names or object/PSCustomObject property names to include, in display order.
        # If not supplied then all Hashtable keys or all object properties will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [Alias('Properties')]
        [AllowNull()]
        [System.String[]] $Columns = $null,

        ## Column widths as percentages. Total should not exceed 100.
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.UInt16[]] $ColumnWidths,

        # Array of custom table header strings in display order.
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [AllowNull()]
        [System.String[]] $Headers = $null,

        ## Table style
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Style = $pscriboDocument.DefaultTableStyle,

        ## Table width (%), 0 = Autofit
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,100)]
        [System.UInt16] $Width = 100,

        ## Indent table
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.UInt16] $Tabs,

        ## List view
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'InputObjectList')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'HashtableList')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'InputObjectListKey')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'HashtableListKey')]
        [System.Management.Automation.SwitchParameter] $List,

        ## Combine list view based upon the specified key property name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'InputObjectListKey')]
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'HashtableListKey')]
        [System.String] $Key,

        ## Hide the key name in the table output. Note: Key name will always be displayed in text output.
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'InputObjectListKey')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'HashtableListKey')]
        [System.Management.Automation.SwitchParameter] $HideKey,

        ## Table caption
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Caption
    )
    begin
    {
        Write-Debug ('Using parameter set "{0}".' -f $PSCmdlet.ParameterSetName)
        [System.Collections.ArrayList] $rows = New-Object -TypeName System.Collections.ArrayList
        Write-PScriboMessage -Message ($localized.ProcessingTable -f $Name)

        if ($Headers -and (-not $Columns))
        {
            Write-PScriboMessage -Message $localized.TableHeadersWithNoColumnsWarning -IsWarning
            $Headers = $Columns
        }
        elseif (($null -ne $Columns) -and ($null -ne $Headers))
        {
            ## Check the number of -Headers matches the number of -Properties
            if ($Headers.Count -ne $Columns.Count)
            {
                Write-PScriboMessage -Message $localized.TableHeadersCountMismatchWarning -IsWarning
                $Headers = $Columns
            }
        }

        if ($ColumnWidths)
        {
            $columnWidthsSum = $ColumnWidths | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            if ($columnWidthsSum -ne 100)
            {
                Write-PScriboMessage -Message ($localized.TableColumnWidthSumWarning -f $columnWidthsSum) -IsWarning
                $ColumnWidths = $null
            }
            elseif ($List -and (-not $Key) -and ($ColumnWidths.Count -ne 2))
            {
                Write-PScriboMessage -Message $localized.ListTableColumnCountWarning -IsWarning
                $ColumnWidths = $null
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'Hashtable') -and (-not $List) -and ($Hashtable[0].Keys.Count -ne $ColumnWidths.Count))
            {
                Write-PScriboMessage -Message $localized.TableColumnWidthMismatchWarning -IsWarning
                $ColumnWidths = $null
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'InputObject') -and (-not $List))
            {
                ## Columns might not have been passed and there is no object in the pipeline here, so check $Columns is an array.
                if (($Columns -is [System.Object[]]) -and ($Columns.Count -ne $ColumnWidths.Count))
                {
                    Write-PScriboMessage -Message $localized.TableColumnWidthMismatchWarning -IsWarning
                    $ColumnWidths = $null
                }
            }
        }
    }
    process
    {
        if ($null -eq $Columns) {
            ## Use all available properties
            if ($PSCmdlet.ParameterSetName -in 'Hashtable','HashtableList','HashtableListKey')
            {
                $Columns = $Hashtable | Select-Object -First 1 -ExpandProperty Keys | Where-Object { $_ -notlike '*__Style' }
            }
            elseif ($PSCmdlet.ParameterSetName -in 'InputObject','InputObjectList','InputObjectListKey')
            {
                ## Pipeline objects are not available in the begin scriptblock
                $object = $InputObject | Select-Object -First 1
                if ($object -is [System.Management.Automation.PSCustomObject])
                {
                    $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name
                }
                elseif ($object -is [System.Collections.Specialized.OrderedDictionary])
                {
                    $Columns = $object.Keys | Where-Object { $_ -notlike '*__Style' }
                }
                else
                {
                    $Columns = Get-Member -InputObject $object -MemberType Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -in 'Hashtable','HashtableList','HashtableListKey')
        {
            foreach ($nestedHashtable in $Hashtable)
            {
                $customObject = New-PScriboTableRow -Hashtable $nestedHashtable
                [ref] $null = $rows.Add($customObject)
            } #end foreach nested hashtable entry
        }
        elseif ($PSCmdlet.ParameterSetName -in 'InputObject','InputObjectList','InputObjectListKey')
        {
            foreach ($object in $InputObject)
            {
                if ($object -is [System.Collections.Specialized.OrderedDictionary])
                {
                    $customObject = New-PScriboTableRow -Hashtable $object
                }
                else
                {
                    $customObject = New-PScriboTableRow -InputObject $object -Properties $Columns -Headers $Headers
                }
                [ref] $null = $rows.Add($customObject)
            }
        }
    }
    end
    {
        if (($PSBoundParameters.ContainsKey('Key')) -and ($PSBoundParameters.ContainsKey('Headers')))
        {
            ## Update the key name to match the rewritten property display name
            $Key = $Headers[$Columns.IndexOf($Key)]
        }

        ## Update the column names as the objects have been have been rewritten with their display names (Headers)
        if ($Headers)
        {
            $Columns = $Headers
        }

        $newPScriboTableParams = @{
            Name         = $Name
            Columns      = $Columns
            ColumnWidths = $ColumnWidths
            Rows         = $rows
            List         = $List
            Style        = $Style
            Width        = $Width
            Tabs         = $Tabs
        }

        if ($PSBoundParameters.ContainsKey('Key'))
        {
            $newPScriboTableParams['ListKey'] = $Key
            $newPScriboTableParams['DisplayListKey'] = -not $HideKey
        }

        if ($PSBoundParameters.ContainsKey('Caption'))
        {
            $newPScriboTableParams['Caption'] = $Caption
        }

        return (New-PScriboTable @newPScriboTableParams)
    }
}

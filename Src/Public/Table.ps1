function Table {
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

    .PARAMETER List

    .PARAMETER Width

    .PARAMETER Tabs


    .EXAMPLE
        Table -Name 'Table 1' -InputObject $(Get-Service) -Columns 'Name','DisplayName','Status' -ColumnWidths 40,20,40



#>
    [CmdletBinding(DefaultParameterSetName = 'InputObject')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Table name/Id
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Name = ([System.Guid]::NewGuid().ToString()),

        # Array of objects
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
        [Alias('CustomObject','Object')]
        [ValidateNotNullOrEmpty()]
        [System.Object[]] $InputObject,

        # Array of Hashtables
        [Parameter(Mandatory, ParameterSetName = 'Hashtable')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Specialized.OrderedDictionary[]] $Hashtable,

        # Array of Hashtable key names or Object/PSCustomObject property names to include, in display order.
        # If not supplied then all Hashtable keys or all PSCustomObject properties will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'InputObject')]
        [Parameter(ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'Hashtable')]
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
        [System.String] $Style = 'TableDefault',

        # List view (no headers)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $List,

        ## Table width (%), 0 = Autofit
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,100)]
        [System.UInt16] $Width = 100,

        ## Indent table
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.UInt16] $Tabs
    ) #end param
    begin {

        <#! Table.Internal.ps1 !#>

        Write-Debug ('Using parameter set "{0}".' -f $PSCmdlet.ParameterSetName);
        [System.Collections.ArrayList] $rows = New-Object -TypeName System.Collections.ArrayList;
        WriteLog -Message ($localized.ProcessingTable -f $Name);

        if ($Headers -and (-not $Columns)) {
            WriteLog -Message $localized.TableHeadersWithNoColumnsWarning -IsWarning;
            $Headers = $Columns;
        } #end if
        elseif (($null -ne $Columns) -and ($null -ne $Headers)) {
            ## Check the number of -Headers matches the number of -Properties
            if ($Headers.Count -ne $Columns.Count) {
                WriteLog -Message $localized.TableHeadersCountMismatchWarning -IsWarning;
                $Headers = $Columns;
            }
        } #end if

        if ($ColumnWidths) {
            $columnWidthsSum = $ColumnWidths | Measure-Object -Sum | Select-Object -ExpandProperty Sum;
            if ($columnWidthsSum -ne 100) {
                WriteLog -Message ($localized.TableColumnWidthSumWarning -f $columnWidthsSum) -IsWarning;
                $ColumnWidths = $null;
            }
            elseif ($List -and $ColumnWidths.Count -ne 2) {
                WriteLog -Message $localized.ListTableColumnCountWarning -IsWarning;
                $ColumnWidths = $null;
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'Hashtable') -and (-not $List) -and ($Hashtable[0].Keys.Count -ne $ColumnWidths.Count)) {
                WriteLog -Message $localized.TableColumnWidthMismatchWarning -IsWarning;
                $ColumnWidths = $null;
            }
            elseif (($PSCmdlet.ParameterSetName -eq 'InputObject') -and (-not $List)) {
                ## Columns might not have been passed and there is no object in the pipeline here, so check $Columns is an array.
                if (($Columns -is [System.Object[]]) -and ($Columns.Count -ne $ColumnWidths.Count)) {
                    WriteLog -Message $localized.TableColumnWidthMismatchWarning -IsWarning;
                    $ColumnWidths = $null;
                }
            }
        } #end if columnwidths

    } #end begin
    process {

        if ($null -eq $Columns) {
            ## Use all available properties
            switch ($PSCmdlet.ParameterSetName) {
                'Hashtable' {
                    $Columns = $Hashtable | Select-Object -First 1 -ExpandProperty Keys | Where-Object { $_ -notlike '*__Style' };
                }
                Default {
                    ## Pipeline objects are not available in the begin scriptblock
                    $object = $InputObject | Select-Object -First 1;
                    if ($object -is [System.Management.Automation.PSCustomObject]) {
                        $Columns = $object.PSObject.Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name;
                    }
                    else {
                        $Columns = Get-Member -InputObject $object -MemberType Properties | Where-Object Name -notlike '*__Style' | Select-Object -ExpandProperty Name;
                    }
                } #end default
            } #end switch parametersetname
        } # end if not columns
        switch ($PSCmdlet.ParameterSetName) {
            'Hashtable' {
                foreach ($nestedHashtable in $Hashtable) {
                    $customObject = New-PScriboTableRow -Hashtable $nestedHashtable;
                    [ref] $null = $rows.Add($customObject);
                } #end foreach nested hashtable entry
            } #end hashtable
            Default {
                foreach ($object in $InputObject) {
                    $customObject = New-PScriboTableRow -InputObject $object -Properties $Columns -Headers $Headers;
                    [ref] $null = $rows.Add($customObject);
                } #end foreach inputobject
            } #end default
        } #end switch

    } #end process
    end {

        ## Reset the column names as the object have been rewritten with their headers
        if ($Headers) { $Columns = $Headers; }
        $table = @{
            Name = $Name;
            Columns = $Columns;
            ColumnWidths = $ColumnWidths;
            Rows = $rows;
            List = $List;
            Style = $Style;
            Width = $Width;
            Tabs = $Tabs;
        }
        return (New-PScriboTable @table);

    } #end end
} #end function Table

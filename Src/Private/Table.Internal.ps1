        #region Table Private Functions

        function New-PScriboTable {
        <#
            .SYNOPSIS
                Initializes a new PScribo table object.

            .PARAMETER Name

            .PARAMETER Columns

            .PARAMETER ColumnWidths

            .PARAMETER Rows

            .PARAMETER Style

            .PARAMETER Width

            .PARAMETER List

            .PARAMETER Width

            .PARAMETER Tabs

            .PARAMETER Caption

            .PARAMETER CaptionTop

            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param
            (
                ## Table name/Id
                [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name = ([System.Guid]::NewGuid().ToString()),

                ## Table columns/display order
                [Parameter(Mandatory)]
                [AllowNull()]
                [System.String[]] $Columns,

                ## Table columns widths
                [Parameter(Mandatory)]
                [AllowNull()]
                [System.UInt16[]] $ColumnWidths,

                ## Collection of PScriboTableObjects for table rows
                [Parameter(Mandatory)]
                [ValidateNotNull()]
                [System.Collections.ArrayList] $Rows,

                ## Table style
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Style,

                # List view
                [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
                [System.Management.Automation.SwitchParameter] $List,

                ## Combine list view based upon the specified key/property name
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'List')]
                [System.String] $ListKey = $null,

                ## Table width (%), 0 = Autofit
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateRange(0, 100)]
                [System.UInt16] $Width = 100,

                ## Indent table
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateRange(0, 10)]
                [System.UInt16] $Tabs,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String] $Caption,

                # Display caption above table.
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $CaptionTop
            )
            process
            {
                $typeName = 'PScribo.Table';
                $pscriboDocument.Properties['Tables']++;
                $pscriboTable = [PSCustomObject] @{
                    Id              = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                    Name            = $Name;
                    Number          = $pscriboDocument.Properties['Tables']
                    Type            = $typeName;
                    # Headers        = $Headers; ## Headers are stored as they may be required when formatting output, i.e. Word tables
                    Columns         = $Columns;
                    ColumnWidths    = $ColumnWidths;
                    Rows            = $Rows;
                    Style           = $Style;
                    Width           = $Width;
                    Tabs            = $Tabs;
                    HasCaption      = $PSBoundParameters.ContainsKey('Caption')
                    IsCaptionBottom = (-not $CaptionTop)
                    Caption         = $Caption
                    IsList          = $List;
                    IsKeyedList     = $PSBoundParameters.ContainsKey('ListKey')
                    ListKey         = $ListKey
                }
                return $pscriboTable;
            }
        }


        function New-PScriboTableRow {
        <#
            .SYNOPSIS
                Defines a new PScribo document table row from an object or hashtable.

            .PARAMETER InputObject

            .PARAMETER Properties

            .PARAMETER Headers

            .PARAMETER Hashtable

            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding(DefaultParameterSetName = 'InputObject')]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param
            (
                ## PSCustomObject to create PScribo table row
                [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
                [System.Object] $InputObject,

                ## PSCutomObject properties to include in the table row
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'InputObject')]
                [AllowNull()]
                [System.String[]] $Properties,

                # Custom table header strings (in Display Order). Used for property names.
                [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'InputObject')]
                [AllowNull()]
                [System.String[]] $Headers = $null,

                ## Array of ordered dictionaries (hashtables) to create PScribo table row
                [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Hashtable')]
                [System.Collections.Specialized.OrderedDictionary] $Hashtable
            )
            begin
            {
                Write-Debug ('Using parameter set "{0}.' -f $PSCmdlet.ParameterSetName);
            }
            process
            {
                switch ($PSCmdlet.ParameterSetName)
                {
                    'Hashtable' {

                        if (-not $Hashtable.Contains('__Style'))
                        {
                            $Hashtable['__Style'] = $null;
                        }
                        ## Create and return custom object from hashtable
                        $psCustomObject = [PSCustomObject] $Hashtable
                        return $psCustomObject;
                    } #end Hashtable

                    Default {

                        $objectProperties = [Ordered] @{ };
                        if ($Properties -notcontains '__Style') { $Properties += '__Style'; }
                        ## Build up hashtable of required property names
                        for ($i = 0; $i -lt $Properties.Count; $i++)
                        {
                            $propertyName = $Properties[$i];
                            $propertyStyleName = '{0}__Style' -f $propertyName;
                            if ($InputObject.PSObject.Properties[$propertyStyleName])
                            {
                                if ($Headers)
                                {
                                    ## Rename the style property to match the header
                                    $headerStyleName = '{0}__Style' -f $Headers[$i];
                                    $objectProperties[$headerStyleName] = $InputObject.$propertyStyleName;
                                }
                                else
                                {
                                    $objectProperties[$propertyStyleName] = $InputObject.$propertyStyleName;
                                }
                            }
                            if ($Headers -and $PropertyName -notlike '*__Style')
                            {
                                if ($InputObject.PSObject.Properties[$propertyName])
                                {
                                    $objectProperties[$Headers[$i]] = $InputObject.$propertyName;
                                }
                            }
                            else
                            {
                                if ($InputObject.PSObject.Properties[$propertyName])
                                {
                                    $objectProperties[$propertyName] = $InputObject.$propertyName;
                                }
                                else
                                {
                                    $objectProperties[$propertyName] = $null;
                                }
                            }
                        } #end foreach property

                        ## Create and return custom object
                        return ([PSCustomObject] $objectProperties);
                    }
                }
            }
        }

        function New-PScriboFormattedTable {
        <#
            .SYNOPSIS
                Creates a formatted table with row and column styling for plugin output/rendering.

            .NOTES
                Maintains backwards compatibility with other plugins that do not require styling/formatting.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Table,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $HasHeaderRow,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $HasHeaderColumn
            )
            process
            {
                $tableStyle = $document.TableStyles[$Table.Style]
                return [PSCustomObject] @{
                    Id              = $Table.Id;
                    Name            = $Table.Name;
                    Number          = $Table.Number;
                    Type            = $Table.Type;
                    ColumnWidths    = $Table.ColumnWidths;
                    Rows            = New-Object -TypeName System.Collections.ArrayList;
                    Width           = $Table.Width;
                    Tabs            = if ($null -eq $Table.Tabs) { 0 } else { $Table.Tabs };
                    PaddingTop      = $tableStyle.PaddingTop;
                    PaddingLeft     = $tableStyle.PaddingLeft;
                    PaddingBottom   = $tableStyle.PaddingBottom;
                    PaddingRight    = $tableStyle.PaddingRight;
                    Align           = $tableStyle.Align;
                    BorderWidth     = $tableStyle.BorderWidth;
                    BorderStyle     = $tableStyle.BorderStyle;
                    BorderColor     = $tableStyle.BorderColor;
                    Style           = $Table.Style;
                    HasHeaderRow    = $HasHeaderRow;  ## First row has header style applied (list/combo list tables)
                    HasHeaderColumn = $HasHeaderColumn;  ## First column has header style applied (list/combo list tables)
                    HasCaption      = $Table.HasCaption;
                    IsCaptionBottom = $Table.IsCaptionBottom; ## Display table caption below
                    Caption         = $Table.Caption;
                    IsList          = $Table.IsList;
                    IsKeyedList     = $Table.IsKeyedList;
                }
            }
        }

        function New-PScriboFormattedTableRow {
        <#
            .SYNOPSIS
                Creates a formatted table row for plugin output/rendering.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.String] $TableStyle,

                [Parameter(ValueFromPipeline)]
                [AllowNull()]
                [System.String] $Style = $null,

                [Parameter(ValueFromPipeline, ParameterSetName = 'Header')]
                [System.Management.Automation.SwitchParameter] $IsHeaderRow,

                [Parameter(ValueFromPipeline, ParameterSetName = 'Row')]
                [System.Management.Automation.SwitchParameter] $IsAlternateRow
            )
            process
            {
                if (-not ([System.String]::IsNullOrEmpty($Style)))
                {
                    ## Use the explictit style
                    $IsStyleInherited = $false
                }
                elseif ($IsHeaderRow)
                {
                    $Style = $document.TableStyles[$TableStyle].HeaderStyle
                    $IsStyleInherited = $true
                }
                elseif ($IsAlternateRow)
                {
                    $Style = $document.TableStyles[$TableStyle].AlternateRowStyle
                    $IsStyleInherited = $true
                }
                else
                {
                    $Style = $document.TableStyles[$TableStyle].RowStyle
                    $IsStyleInherited = $true
                }

                return [PSCustomObject] @{
                    Style            = $Style;
                    IsStyleInherited = $IsStyleInherited;
                    Cells            = New-Object -TypeName System.Collections.ArrayList;
                }
            }
        }

        function New-PScriboFormattedTableRowCell {
        <#
            .SYNOPSIS
                Creates a formatted table cell for plugin output/rendering.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(ValueFromPipeline)]
                [AllowNull()]
                [System.String] $Content,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Uint16] $Width,

                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.String] $Style = $null,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Boolean] $IsAlternateRow
            )
            process
            {
                $isStyleInherited = $true

                if (-not ([System.String]::IsNullOrEmpty($Style)))
                {
                    ## Use the explictit style
                    $isStyleInherited = $false
                }

                return [PSCustomObject] @{
                    Content          = $Content;
                    Width            = $Width;
                    Style            = $Style;
                    IsStyleInherited = $isStyleInherited;
                }
            }
        }

        function ConvertTo-PScriboPreformattedTable {
        <#
            .SYNOPSIS
                Creates a formatted table based upon table type for plugin output/rendering.

            .NOTES
                Maintains backwards compatibility with other plugins that do not require styling/formatting.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Table
            )
            process
            {
                if ($Table.IsKeyedList)
                {
                    Write-Output -InputObject (ConvertTo-PScriboFormattedKeyedListTable -Table $Table)
                }
                elseif ($Table.IsList)
                {
                    Write-Output -InputObject (ConvertTo-PScriboFormattedListTable -Table $Table)
                }
                else
                {
                    Write-Output -InputObject (ConvertTo-PScriboFormattedTable -Table $Table)
                }
            }
        }

        function ConvertTo-PScriboFormattedTable {
        <#
            .SYNOPSIS
                Creates a formatted standard table for plugin output/rendering.

            .NOTES
                Maintains backwards compatibility with other plugins that do not require styling/formatting.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Table
            )
            begin
            {
                $hasColumnWidths = ($null -ne $Table.ColumnWidths)
            }
            process
            {
                $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderRow

                # Output the header row and header cells
                $headerRow = New-PScriboFormattedTableRow -TableStyle $Table.Style -IsHeaderRow
                for ($h = 0; $h -lt $Table.Columns.Count; $h++)
                {
                    $newPScriboFormattedTableHeaderCellParams = @{
                        Content    = $Table.Columns[$h]
                    }
                    if ($hasColumnWidths)
                    {
                        $newPScriboFormattedTableHeaderCellParams['Width'] = $Table.ColumnWidths[$h]
                    }
                    $cell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableHeaderCellParams
                    $null = $headerRow.Cells.Add($cell)
                }
                $null = $formattedTable.Rows.Add($headerRow)

                ## Output each object row
                for ($r = 0; $r -lt $Table.Rows.Count; $r++)
                {
                    $objectProperties = $Table.Rows[$r].PSObject.Properties

                    $newPScriboFormattedTableRowParams = @{
                        TableStyle = $Table.Style;
                        Style = $Table.Rows[$r].'__Style'
                        IsAlternateRow = ($r % 2 -ne 0 )
                    }
                    $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

                    ## Output object row's cells
                    for ($c = 0; $c -lt $Table.Columns.Count; $c++)
                    {
                        $propertyName = $Table.Columns[$c]
                        $propertyStyleName = '{0}__Style' -f $propertyName;
                        $hasStyleProperty = $objectProperties.Name.Contains($propertyStyleName)

                        $propertyValue = $objectProperties[$propertyName].Value
                        $newPScriboFormattedTableRowCellParams = @{
                            Content = $propertyValue
                        }
                        if ([System.String]::IsNullOrEmpty($propertyValue))
                        {
                            $newPScriboFormattedTableRowCellParams['Content'] = $null
                        }

                        if ($hasColumnWidths)
                        {
                            $newPScriboFormattedTableRowCellParams['Width'] = $Table.ColumnWidths[$c]
                        }
                        if ($hasStyleProperty)
                        {
                            $newPScriboFormattedTableRowCellParams['Style'] = $objectProperties[$propertyStyleName].Value # | Where-Object Name -eq $propertyStyleName | Select-Object -ExpandProperty Value
                        }

                        $cell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowCellParams
                        $null = $row.Cells.Add($cell)
                    }
                    $null = $formattedTable.Rows.Add($row)
                }
                Write-Output -InputObject $formattedTable
            }
        }

        function ConvertTo-PScriboFormattedListTable {
        <#
            .SYNOPSIS
                Creates a formatted list table for plugin output/rendering.

            .NOTES
                Maintains backwards compatibility with other plugins that do not require styling/formatting.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSObject[]])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Table
            )
            begin
            {
                $hasColumnWidths = ($null -ne $Table.ColumnWidths)
            }
            process
            {
                for ($r = 0; $r -lt $Table.Rows.Count; $r++)
                {
                    ## We have one table per object
                    $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderColumn
                    $objectProperties = $Table.Rows[$r].PSObject.Properties

                    for ($c = 0; $c -lt $Table.Columns.Count; $c++)
                    {
                        $column = $Table.Columns[$c]
                        $newPScriboFormattedTableRowParams = @{
                            TableStyle = $Table.Style;
                            Style = $Table.Rows[$r].'__Style'
                            IsAlternateRow = ($c % 2) -ne 0
                        }
                        $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

                        ## Add each property name (as header style)
                        $newPScriboFormattedTableRowHeaderCellParams = @{
                            Content = $column
                        }
                        if ($hasColumnWidths)
                        {
                            $newPScriboFormattedTableRowHeaderCellParams['Width'] = $Table.ColumnWidths[0]
                        }
                        $headerCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
                        $null = $row.Cells.Add($headerCell)

                        ## Add each property value
                        $propertyValue = $objectProperties[$column].Value
                        if ([System.String]::IsNullOrEmpty($propertyValue))
                        {
                            $propertyValue = $null
                        }

                        $newPScriboFormattedTableRowValueCellParams = @{
                            Content = $propertyValue
                        }

                        $propertyStyleName = '{0}__Style' -f $column
                        $hasStyleProperty = $objectProperties.Name.Contains($propertyStyleName)
                        if ($hasStyleProperty)
                        {
                            $newPScriboFormattedTableRowValueCellParams['Style'] = $objectProperties[$propertyStyleName].Value
                        }
                        if ($hasColumnWidths)
                        {
                            $newPScriboFormattedTableRowValueCellParams['Width'] = $Table.ColumnWidths[1]
                        }
                        $valueCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowValueCellParams
                        $null = $row.Cells.Add($valueCell)

                        $null = $formattedTable.Rows.Add($row)
                    }
                    Write-Output -InputObject $formattedTable
                }
            }
        }

        function ConvertTo-PScriboFormattedKeyedListTable {
        <#
            .SYNOPSIS
                Creates a formatted keyed list table (a key'd column per object) for plugin output/rendering.

            .NOTES
                Maintains backwards compatibility with other plugins that do not require styling/formatting.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSObject[]])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Table
            )
            begin
            {
                $hasColumnWidths = ($null -ne $Table.ColumnWidths)
                $objectKey = $Table.ListKey
            }
            process
            {
                $formattedTable = New-PScriboFormattedTable -Table $Table -HasHeaderRow -HasHeaderColumn

                ## Output the header cells
                $headerRow = New-PScriboFormattedTableRow -TableStyle $Table.Style -IsHeaderRow

                ## The top-left cell is always empty
                $newPScriboFormattedTableRowHeaderCellParams = @{
                    Content = ' '
                }
                $blankHeaderCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
                $null = $headerRow.Cells.Add($blankHeaderCell)

                ## Add all the object key values
                for ($o = 0; $o -lt $Table.Rows.Count; $o++)
                {
                    $newPScriboFormattedTableRowHeaderCellParams = @{
                        Content = $Table.Rows[$o].PSObject.Properties[$objectKey].Value
                    }
                    $objectKeyCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowHeaderCellParams
                    $null = $headerRow.Cells.Add($objectKeyCell)
                }
                $null = $formattedTable.Rows.Add($headerRow)

                $isAlternateRow = $false
                ## Output remaining object properties (one property per row)
                foreach ($column in $Table.Columns)
                {
                    if ((-not $column.EndsWith('__Style', 'CurrentCultureIgnoreCase')) -and
                        ($column -ne $objectKey))
                    {
                        ## Add the object property column
                        $newPScriboFormattedTableRowParams = @{
                            TableStyle = $Table.Style;
                            IsAlternateRow = $isAlternateRow
                        }
                        $row = New-PScriboFormattedTableRow @newPScriboFormattedTableRowParams

                        ## Output the column header cell (property name) as header style
                        $newPScriboFormattedTableColumnCellParams = @{
                            Content = $column
                        }
                        if ($hasColumnWidths) {
                            $newPScriboFormattedTableColumnCellParams['Width'] = $Table.ColumnWidths[0]
                        }
                        $columnCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableColumnCellParams
                        $null = $row.Cells.Add($columnCell)

                        ## Add the property value for all other objects
                        for ($o = 0; $o -lt $Table.Rows.Count; $o++)
                        {
                            $propertyValue = $Table.Rows[$o].PSObject.Properties[$column].Value
                            $newPScriboFormattedTableRowValueCellParams = @{
                                Content = $propertyValue
                            }
                            if ([System.String]::IsNullOrEmpty($propertyValue)) {
                                $newPScriboFormattedTableRowValueCellParams['Content'] = $null
                            }
                            $propertyStyleName = '{0}__Style' -f $column
                            $hasStyleProperty = $Table.Rows[$o].PSObject.Properties.Name.Contains($propertyStyleName)
                            if ($hasStyleProperty) {
                                $newPScriboFormattedTableRowValueCellParams['Style'] = $Table.Rows[$o].PSObject.Properties[$propertyStyleName].Value
                            }
                            if ($hasColumnWidths) {
                                $newPScriboFormattedTableRowValueCellParams['Width'] = $Table.ColumnWidths[$o+1]
                            }
                            $valueCell = New-PScriboFormattedTableRowCell @newPScriboFormattedTableRowValueCellParams
                            $null = $row.Cells.Add($valueCell)
                        }

                        $null = $formattedTable.Rows.Add($row)
                        $isAlternateRow = -not $isAlternateRow
                    }
                }
                Write-Output -InputObject $formattedTable
            }
        }

        #endregion Table Private Functions

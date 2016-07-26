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

        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
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

                ## List view
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
            process {

                $typeName = 'PScribo.Table';
                $pscriboDocument.Properties['Tables']++;
                $pscriboTable = [PSCustomObject] @{
                    Id = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                    Name = $Name;
                    Type = $typeName;
                    # Headers = $Headers; ## Headers are stored as they may be required when formatting output, i.e. Word tables
                    Columns = $Columns;
                    ColumnWidths = $ColumnWidths;
                    Rows = $Rows;
                    List = $List;
                    Style = $Style;
                    Width = $Width;
                    Tabs = $Tabs;
                }
                return $pscriboTable;

            } #end process
        } #end function new-pscribotable


        function New-PScriboTableRow {
        <#
            .SYNOPSIS
                Defines a new PScribo document table row from an object or hashtable.

            .PARAMETER InputObject

            .PARAMETER Properties

            .PARAMETER Headers

            .PARAMETER Hashtable

        #>
            [CmdletBinding(DefaultParameterSetName='InputObject')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PSCustomObject to create PScribo table row
                [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
                [ValidateNotNull()]
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
                [AllowNull()]
                [System.Collections.Specialized.OrderedDictionary] $Hashtable
            )
            begin {

                Write-Debug ('Using parameter set "{0}.' -f $PSCmdlet.ParameterSetName);

            } #end begin
            process {

                switch ($PSCmdlet.ParameterSetName) {
                    'Hashtable'{
                        if (-not $Hashtable.Contains('__Style')) {
                            $Hashtable['__Style'] = $null;
                        }
                        ## Create and return custom object from hashtable
                        return ([PSCustomObject] $Hashtable);
                    } #end Hashtable
                    Default {
                        $objectProperties = [Ordered] @{ };
                        if ($Properties -notcontains '__Style') { $Properties += '__Style'; }
                        ## Build up hashtable of required property names
                        for ($i = 0; $i -lt $Properties.Count; $i++) {
                            $propertyName = $Properties[$i];
                            $propertyStyleName = '{0}__Style' -f $propertyName;
                            if ($InputObject.PSObject.Properties[$propertyStyleName]) {
                                if ($Headers) {
                                    ## Rename the style property to match the header
                                    $headerStyleName = '{0}__Style' -f $Headers[$i];
                                    $objectProperties[$headerStyleName] = $InputObject.$propertyStyleName;
                                }
                                else {
                                    $objectProperties[$propertyStyleName] = $InputObject.$propertyStyleName;
                                }
                            }
                            if ($Headers -and $PropertyName -notlike '*__Style') {
                                if ($InputObject.PSObject.Properties[$propertyName]) {
                                    $objectProperties[$Headers[$i]] = $InputObject.$propertyName;
                                }
                            }
                            else {
                                if ($InputObject.PSObject.Properties[$propertyName]) {
                                    $objectProperties[$propertyName] = $InputObject.$propertyName;
                                }
                                else {
                                    $objectProperties[$propertyName] = $null;
                                }
                            }
                        } #end for
                        ## Create and return custom object
                        return ([PSCustomObject] $objectProperties);
                    } #end Default
                } #end switch

            } #end process
        } #end function New-PScriboTableRow

        #endregion Table Private Functions

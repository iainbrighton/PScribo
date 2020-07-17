function New-PScriboTableRow
{
<#
    .SYNOPSIS
        Defines a new PScribo document table row from an object or hashtable.

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
        Write-Debug ('Using parameter set "{0}.' -f $PSCmdlet.ParameterSetName)
    }
    process
    {
        switch ($PSCmdlet.ParameterSetName)
        {
            'Hashtable'
            {
                if (-not $Hashtable.Contains('__Style'))
                {
                    $Hashtable['__Style'] = $null
                }
                ## Create and return custom object from hashtable
                $psCustomObject = [PSCustomObject] $Hashtable
                return $psCustomObject
            }
            Default
            {
                $objectProperties = [Ordered] @{ }
                if ($Properties -notcontains '__Style')
                {
                    $Properties += '__Style'
                }

                ## Build up hashtable of required property names
                for ($i = 0; $i -lt $Properties.Count; $i++)
                {
                    $propertyName = $Properties[$i]
                    $propertyStyleName = '{0}__Style' -f $propertyName
                    if ($InputObject.PSObject.Properties[$propertyStyleName])
                    {
                        if ($Headers)
                        {
                            ## Rename the style property to match the header
                            $headerStyleName = '{0}__Style' -f $Headers[$i]
                            $objectProperties[$headerStyleName] = $InputObject.$propertyStyleName
                        }
                        else
                        {
                            $objectProperties[$propertyStyleName] = $InputObject.$propertyStyleName
                        }
                    }
                    if ($Headers -and $PropertyName -notlike '*__Style')
                    {
                        if ($InputObject.PSObject.Properties[$propertyName])
                        {
                            $objectProperties[$Headers[$i]] = $InputObject.$propertyName
                        }
                    }
                    else
                    {
                        if ($InputObject.PSObject.Properties[$propertyName])
                        {
                            $objectProperties[$propertyName] = $InputObject.$propertyName
                        }
                        else
                        {
                            $objectProperties[$propertyName] = $null
                        }
                    }
                }

                ## Create and return custom object
                return ([PSCustomObject] $objectProperties)
            }
        }
    }
}

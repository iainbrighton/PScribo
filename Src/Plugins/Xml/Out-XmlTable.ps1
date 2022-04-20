function Out-XmlTable
{
<#
    .SYNOPSIS
        Output formatted Xml table.
#>
    [CmdletBinding()]
    param
    (
        ## PScribo table object
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        $tableId = ($Table.Id -replace '[^a-z0-9-_\.]','').ToLower()
        $tableElement = $element.AppendChild($xmlDocument.CreateElement($tableId))
        [ref] $null = $tableElement.SetAttribute('name', $Table.Name)

        foreach ($row in $Table.Rows)
        {
            $groupElement = $tableElement.AppendChild($xmlDocument.CreateElement('group'))
            foreach ($property in $row.PSObject.Properties)
            {
                if (-not ($property.Name).EndsWith('__Style', 'CurrentCultureIgnoreCase'))
                {
                    $propertyId = ($property.Name -replace '[^a-z0-9-_\.]','').ToLower()
                    $rowElement = $groupElement.AppendChild($xmlDocument.CreateElement($propertyId))
                    ## Only add the Name attribute if there's a difference
                    if ($property.Name -ne $propertyId)
                    {
                        [ref] $null = $rowElement.SetAttribute('name', $property.Name)
                    }
                    [ref] $null = $rowElement.AppendChild($xmlDocument.CreateTextNode($row.($property.Name)))
                }
            }
        }

        return $tableElement
    }
}

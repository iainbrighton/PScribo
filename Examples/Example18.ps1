Import-Module PScribo -Force;

$example18 = Document -Name 'PScribo Example 18' {
    <#
        If you need to list an object's or hashtable's values in a list format, you
        can specify the -List parameter on the 'Table' cmdlet. Rather than creating
        a table column for each property, it will create a two-column table, with a
        row for each property instead.

        This is useful if not all properties will fit across as page (like Services).
        However, if multiple objects are encountered, PScribo will create a separate
        two-column table for each object.

        The following example will create a list table, detailing every property,
        for every service regiestered on the local machine.
    #>
    Get-Service | Table -List
}
$example18 | Export-Document -Format Html -Path ~\Desktop

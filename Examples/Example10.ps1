Import-Module PScribo -Force;

$example10 = Document -Name 'PScribo Example 10' {
    <#
       PScribo supports insertion of tables into a document. In it's simplest form, a
       collection of objects can be added to a document by using the 'Table' cmdlet.

       The following examples creates a table of all properties on all services.

       NOTE: Due to the number of properties, the table will not render in any
       meaningful way and WILL overflow the containers!
    #>
    Table -InputObject (Get-Service)
}
$example10 | Export-Document -Format Html -Path ~\Desktop

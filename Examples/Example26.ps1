[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example26 = Document -Name 'PScribo Example 26' {

    <#
        There may be occassions when you need to highlight rows within a table. Before
        styling can be applied to a row (other than the table's default row and alternate
        row styles), it must first be defined.

        The following defines a custom style that we will apply to particular table rows
        later.

        NOTE: You can combine custom row styling with custom table styles!
    #>
    Style -Name StoppedService -Color White -BackgroundColor Firebrick

    <#
        To apply a style to a particular table row, an additional property called '__Style'
        needs to be added to each object that we need apply an alternate style to. The
        value of this string property is the name of the custom style to apply.

        How the '__Style' property is added to an individual object or objects within a
        collection does not matter. However, PScribo provides the Set-Style cmdlet to
        easily add the property to one or more objects.

        The following retrieves all local services and then enumerates if the service is
        stopped and sets the style (of the entire row) for those objects to the custom
        'StoppedService' style.

        NOTE: We need to retrieve all the services into a collection and then enumerate
              them - otherwise we would only have stopped services to display!
    #>
    $services = Get-Service
    $services | Where-Object { $_.Status -ne 'Running' } | Set-Style -Style 'StoppedService'

    <#
        Once the '__Style' property has been added to the relevant objects (Services) in
        the collection, we can then create the table.
    #>
    Table -InputObject $services -Columns Name,DisplayName,Status -Headers 'Name','Display Name','State'
}
$example26 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

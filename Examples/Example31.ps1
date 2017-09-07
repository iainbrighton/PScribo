param ([System.Management.Automation.SwitchParameter] $PassThru)

# Import-Module PScribo -Force;

$example31 = Document -Name 'PScribo Example 31' {
    <#
        When sending PScribo output as a HTML email body, Outlook does not necessarily
        display the page styling/look-and-feel correctly. This can be suppressed by
        passing an -Options hashtable with the 'NoPageLayoutStyle' key to the
        Export-Document function.
    #>
    DocumentOption -Margin 18
    Section -Style 'Heading1' -Name 'Embedded Image' {

        Image -Path "$PSScriptRoot\Example31.jpg" -Height 160 -Width 160

        Paragraph 'Image Attribution: https://cdn.pixabay.com/photo/2014/08/26/19/20/document-428334_640.jpg'
    }
}
$example31 | Export-Document -Format Word -Path ~\Desktop -PassThru:$PassThru

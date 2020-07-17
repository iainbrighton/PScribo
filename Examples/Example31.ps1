[CmdletBinding()]
param (
    [System.String[]] $Format = 'Word',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

$example31 = Document -Name 'PScribo Example 31' {

    <#
        Images can be added from a local file, the web or included as a Base64 encoded string. Images are output in
        the Word, Html and XML formats. Text output includes the image's alternative text description but doesn't
        include the binary content. By default, when inserting an image the alternative text is set the image's Uri.

        When images are added to the PScribo document, the images' file contents are stored interanlly to the document.
        This means that external images (files or web content) are only required when generating the document. If you
        have a PScribo document object, access to the files/internet is not required.
    #>
    Section -Style 'Heading1' -Name 'Base64 encoded image at native size' {

        <#
            Embedding images as Base64 encoded strings within the document has one major advantage - there is no
            reliance on external image files or access to the internet.

            NOTE: When inserting Base64 encoded images you must include the alternative text.

            You can convert an existing image file to Base64 using the [System.Convert]::ToBase64() static method, i.e.
            [System.Convert]::ToBase64String((Get-Content -Path .\ExampleHtmlOutput.png -Encoding Byte))
        #>
        Image -Text 'Soccer man' -Base64 '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'
        Paragraph 'Image Attribution: https://icons8.com/icon/42919/soccer' -Size 8 -Italic
    }

    Section -Style 'Heading1' -Name 'Image from file with defined dimensions' {

        <#
            Images are displayed at their native resolution. If you want to resize an image, specify the required height
            and width parameters. Remember to make sure you maintain the image's aspect ratio!
        #>
        Image -Path "$PSScriptRoot\Example31.jpg" -Height 160 -Width 160
        Paragraph 'Image Attribution: https://cdn.pixabay.com/photo/2014/08/26/19/20/document-428334_640.jpg' -Size 8 -Italic
    }

    Section -Style 'Heading1' -Name 'Centered image from web scaled proportionally by 50%' {

        <#
            Images can also be resized by a percentage. Using '-Percent 100' will display the image at it's native
            resolution.

            For example, if you wanted to reduce the image's size to 75% of its original size (whilst maintaining the
            aspect ratio), you would specify '-Percent 75'. If you wanted to increase the image's size by 25% (whilst
            mainaining the aspect ration), you would specify '-Percent 125'.
        #>
        Image -Uri 'https://cdn.pixabay.com/photo/2014/08/26/19/20/document-428334_640.jpg' -Percent 50 -Align Center
        Paragraph 'Image Attribution: https://cdn.pixabay.com/photo/2014/08/26/19/20/document-428334_640.jpg' -Size 8 -Italic
    }

}
$example31 | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

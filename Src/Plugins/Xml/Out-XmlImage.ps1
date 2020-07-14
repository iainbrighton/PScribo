function Out-XmlImage
{
<#
    .SYNOPSIS
        Output embedded Xml image.
#>
    [CmdletBinding()]
    param
    (
        ## PScribo Image object
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Image
    )
    process
    {
        $imageElement = $xmlDocument.CreateElement('image')
        [ref] $null = $imageElement.SetAttribute('text', $Image.Text)
        [ref] $null = $imageElement.SetAttribute('mimeType', $Image.MimeType)
        $imageBase64 = [System.Convert]::ToBase64String($Image.Bytes)
        [ref] $null = $imageElement.AppendChild($xmlDocument.CreateTextNode($imageBase64))

        return $imageElement
    }
}

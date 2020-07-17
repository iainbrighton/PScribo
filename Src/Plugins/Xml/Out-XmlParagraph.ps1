function Out-XmlParagraph
{
<#
    .SYNOPSIS
        Output formatted Xml paragraph run.
#>
    [CmdletBinding()]
    param
    (
        ## PScribo paragraph object
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Paragraph
    )
    process
    {
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        foreach ($paragraphRun in $Paragraph.Sections)
        {

            $text = Resolve-PScriboToken -InputObject $paragraphRun.Text
            [ref] $null = $paragraphBuilder.Append($text)

            if (($paragraphRun.IsParagraphRunEnd -eq $false) -and
                ($paragraphRun.NoSpace -eq $false))
            {
                [ref] $null = $paragraphBuilder.Append(' ')
            }
        }

        $paragraphElement = $xmlDocument.CreateElement('paragraph')
        [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($paragraphBuilder.ToString()))

        return $paragraphElement
    }
}

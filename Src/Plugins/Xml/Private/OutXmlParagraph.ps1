function OutXmlParagraph
{
<#
    .SYNOPSIS
        Output formatted Xml paragraph.
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
        if (-not ([string]::IsNullOrEmpty($Paragraph.Value)))
        {
            ## Value override specified
            $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
            $paragraphElement = $xmlDocument.CreateElement($paragraphId);
            [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Value));
        }
        elseif ([string]::IsNullOrEmpty($Paragraph.Text))
        {
            ## No Id/Name specified, therefore insert as a comment
            $paragraphElement = $xmlDocument.CreateComment((' {0} ' -f $Paragraph.Id));
        }
        else
        {
            ## Create an element with the Id/Name
            $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
            $paragraphElement = $xmlDocument.CreateElement($paragraphId);
            [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Text));
        }
        return $paragraphElement;
    }
}

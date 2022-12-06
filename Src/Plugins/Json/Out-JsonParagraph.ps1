function Out-JsonParagraph
{
<#
    .SYNOPSIS
        Output formatted paragraph run.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Paragraph
    )
    begin
    {
        ## Initializing string object
        [System.Text.StringBuilder] $paragraphBuilder = New-Object -TypeName 'System.Text.StringBuilder'
    }
    process
    {
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


        return $paragraphBuilder.ToString()
    }
}

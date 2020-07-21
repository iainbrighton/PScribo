function Out-MarkdownTOC
{
<#
    .SYNOPSIS
        Output formatted Table of Contents
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $TOC
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboMarkdownOption
        }
    }
    process
    {
        $tocBuilder = New-Object -TypeName System.Text.StringBuilder
        $tocName = '# {0}' -f $TOC.Name
        [ref] $null = $tocBuilder.AppendLine($tocName).AppendLine()

        if ($Options.ContainsKey('EnableSectionNumbering'))
        {
            foreach ($tocEntry in $Document.TOC)
            {
                $tocLink = $tocEntry.Name.Replace(' ','-').ToLower()
                [ref] $null = $tocBuilder.AppendFormat('[{0} {1}](#{0}-{2}) <br />', $tocEntry.Number, $tocEntry.Name, $tocLink).AppendLine()
            }
        }
        else
        {
            foreach ($tocEntry in $Document.TOC)
            {
                $tocLink = $tocEntry.Name.Replace(' ','-').ToLower()
                [ref] $null = $tocBuilder.AppendFormat('[{0}](#{1}) <br />', $tocEntry.Name, $tocLink).AppendLine()
            }
        }
        [ref] $null = $tocBuilder.AppendLine()

        return $tocBuilder.ToString()
    }
}

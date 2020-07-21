function Out-MarkdownBlankLine
{
<#
    .SYNOPSIS
        Output formatted Markdown blankline.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $BlankLine
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
        $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder
        for ($i = 0; $i -lt $BlankLine.LineCount; $i++)
        {
            if ($options.RenderBlankLine)
            {
                [ref] $null = $blankLineBuilder.Append('<br />')
            }
            [ref] $null = $blankLineBuilder.AppendLine()
        }
        return $blankLineBuilder.ToString()
    }
}

function Out-TextBlankLine
{
<#
    .SYNOPSIS
        Output formatted text blankline.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $BlankLine
    )
    process
    {
        $blankLineBuilder = New-Object -TypeName System.Text.StringBuilder
        for ($i = 0; $i -lt $BlankLine.LineCount; $i++)
        {
            [ref] $null = $blankLineBuilder.AppendLine()
        }
        return $blankLineBuilder.ToString()
    }
}

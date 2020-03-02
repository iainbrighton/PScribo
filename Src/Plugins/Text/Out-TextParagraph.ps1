function Out-TextParagraph
{
<#
    .SYNOPSIS
        Output formatted paragraph text.
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
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboTextOption;
        }
    }
    process
    {
        $padding = ''.PadRight(($Paragraph.Tabs * 4), ' ')
        if ([string]::IsNullOrEmpty($Paragraph.Text))
        {
            $text = "$padding$($Paragraph.Id)"
        }
        else
        {
            $text = "$padding$($Paragraph.Text)"
        }
        $formattedText = Out-StringWrap -InputObject $text -Width $Options.TextWidth
        if ($Paragraph.NewLine)
        {
            $formattedText = '{0}{1}' -f $formattedText, [System.Environment]::NewLine
        }
        return $formattedText
    }
}

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
        $convertToAlignedStringParams = @{
            Width       = $options.TextWidth
            Tabs        = $Paragraph.Tabs
            Align       = 'Left'
            NoNewLine   = -not $Paragraph.NewLine
        }

        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $paragraphStyle = Get-PScriboDocumentStyle -Style $Paragraph.Style
            $convertToAlignedStringParams['Align'] = $paragraphStyle.Align
        }

        if ([string]::IsNullOrEmpty($Paragraph.Text))
        {
            $convertToAlignedStringParams['InputObject'] = $Paragraph.Id
        }
        else
        {
            $convertToAlignedStringParams['InputObject'] = $Paragraph.Text
        }

        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

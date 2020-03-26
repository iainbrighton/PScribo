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
        [System.Management.Automation.PSObject] $Paragraph,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ParseToken
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

        $text = $Paragraph.Text
        if ([string]::IsNullOrEmpty($text))
        {
            $text = $Paragraph.Id
        }

        if ($ParseToken)
        {
            $text = Resolve-PScriboToken -InputObject $text
        }

        $convertToAlignedStringParams['InputObject'] = $text

        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

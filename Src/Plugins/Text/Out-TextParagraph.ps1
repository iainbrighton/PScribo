function Out-TextParagraph
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
        }

        if (-not [System.String]::IsNullOrEmpty($Paragraph.Style))
        {
            $paragraphStyle = Get-PScriboDocumentStyle -Style $Paragraph.Style
            $convertToAlignedStringParams['Align'] = $paragraphStyle.Align
        }

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

        $convertToAlignedStringParams['InputObject'] = $paragraphBuilder.ToString()
        return (ConvertTo-AlignedString @convertToAlignedStringParams)
    }
}

function Text
{
<#
    .SYNOPSIS
        Paragraphs can be made up of text "blocks" (runs) concatenated together to
        form a body of text. Separating the text into runs permits alteration
        of styling element within a single paragraph.

    .NOTES
        The "Text" block command can only be used within a Paragraph -ScriptBlock { }
#>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [AllowEmptyString()]
        [System.String] $Text,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Style')]
        [System.String] $Style,

        ## No space applied between this text block and the next
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoSpace,

        ## Override the bold style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Bold,

        ## Override the italic style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Italic,

        ## Override the underline style
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [System.Management.Automation.SwitchParameter] $Underline,

        ## Override the font name(s)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Font,

        ## Override the font size (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [AllowNull()]
        [System.UInt16] $Size = $null,

        ## Override the font color/colour
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Inline')]
        [AllowNull()]
        [System.String] $Color = $null
    )
    begin
    {
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($psCallStack[1].FunctionName -ne 'Paragraph<Process>')
        {
            throw $localized.ParagraphRunRootError
        }
    }
    process
    {
        $paragraphRunDisplayName = $Text
        if ($Text.Length -gt 40)
        {
            $paragraphRunDisplayName = '{0}[..]' -f $Text.Substring(0,36)
        }
        Write-PScriboMessage -Message ($localized.ProcessingParagraphRun -f $paragraphRunDisplayName)
        return (New-PScriboParagraphRun @PSBoundParameters)
    }
}

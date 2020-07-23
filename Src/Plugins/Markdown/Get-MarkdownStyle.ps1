function Get-MarkdownStyle
{
<#
    .SYNOPSIS
        Generates Markdown styling markup.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Bold,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Italic
    )
    process
    {
        if ($Bold -and $Italic)
        {
            return '***'
        }
        elseif ($Bold)
        {
            return '**'
        }
        elseif ($Italic)
        {
            return '_'
        }
    }
}

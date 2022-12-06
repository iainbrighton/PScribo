function Get-JsonTableCaption
{
<#
    .SYNOPSIS
        Generates caption from a PScribo.Table object.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Table
    )
    process
    {
        $tableStyle = Get-PScriboDocumentStyle -TableStyle $Table.Style

        return [PSCustomObject] @{
            Caption = ('{0} {1} {2}' -f $tableStyle.CaptionPrefix, $Table.CaptionNumber, $Table.Caption)
        }
    }
}

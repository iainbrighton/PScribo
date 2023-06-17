function Get-PScriboListItemMaximumLength
{
<#
    .SYNOPSIS
        Renders a List's item numbers to determine the maximum string/render width.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','NumberStyle')]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.PSObject] $NumberStyle
    )
    process
    {
        $List.Items |
            Where-Object { $_.Type -eq 'PScribo.Item' } |
                ForEach-Object {
                    $number = ConvertFrom-NumberStyle -Value $_.Index -NumberStyle $numberStyle
                    Write-Output -InputObject $number.Length
                } |
                    Measure-Object -Maximum |
                        Select-Object -ExpandProperty Maximum
    }
}

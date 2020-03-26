function New-PScriboPageBreak
{
<#
    .SYNOPSIS
        Creates a PScribo page break object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    process
    {
        $typeName = 'PScribo.PageBreak';
        $pscriboDocument.Properties['PageBreaks']++;
        $pscriboDocument.Properties['Pages']++;
        $pscriboPageBreak = [PSCustomObject] @{
            Id   = $Id;
            Type = $typeName;
        }
        return $pscriboPageBreak;
    }
}

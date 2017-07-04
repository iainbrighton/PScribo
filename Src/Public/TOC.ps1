function TOC {
<#
    .SYNOPSIS
        Initializes a new PScribo Table of Contents (TOC) object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        [Parameter(ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name = 'Contents',

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $ClassId = 'TOC'
    )
    begin {

        <#! TOC.Internal.ps1 !#>

    } #end begin
    process {

        WriteLog -Message ($localized.ProcessingTOC -f $Name);
        return (New-PScriboTOC @PSBoundParameters);

    }
} #end function TOC

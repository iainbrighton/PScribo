function Image {
<#
    .SYNOPSIS
        Initializes a new PScribo Image object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Image Id and Xml element name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## Image text. If empty $Name/Id will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [AllowNull()]
        [System.String] $Text = $null,

        ## Output value override, i.e. for Xml elements. If empty $Text will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [AllowNull()]
        [System.String] $Value = $null,
        ## FilePath
        [Parameter()]
        [System.String] $FilePath,
        ## Tab indent
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0
    )
    begin {

        <#! Image.Internal.ps1 !#>

    } #end begin
    process {

        if ($Name.Length -gt 40) { $ImageDisplayName = '{0}[..]' -f $Name.Substring(0,36); }
        else { $ImageDisplayName = $Name; }
        WriteLog -Message ($localized.ProcessingImage -f $ImageDisplayName);
        return (New-PScriboImage @PSBoundParameters);

    } #end process
} #end function Image

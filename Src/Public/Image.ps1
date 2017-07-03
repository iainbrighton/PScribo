function Image 
{
    <#
            .SYNOPSIS
            Initializes a new PScribo Image object.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (

        ## FilePath
        [Parameter(Mandatory,ValueFromPipelineByPropertyName, Position = 0)]
        [System.String] $FilePath,
        ## FilePath will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [AllowNull()]
        [System.String] $Text = $null,
        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [AllowNull()]
        [Int32] $PixelHeight = $null,
        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [AllowNull()]
        [Int32] $PixelWidth = $null
    )
    begin {

        <#! Image.Internal.ps1 !#>

    } #end begin
    process {
        
        if ($Text.Length -gt 40) 
        {
            $ImageDisplayName = '{0}[..]' -f $Text.Substring(0,36)
        }
        else 
        {
            $ImageDisplayName = $Text
        }
        WriteLog -Message ($localized.ProcessingImage -f $ImageDisplayName)
        return (New-PScriboImage @PSBoundParameters)

    } #end process
} #end function Image

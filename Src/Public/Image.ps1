function Image {
    <#
    .SYNOPSIS
        Initializes a new PScribo Image object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## File path
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('FilePath','Uri')]
        [System.String] $Path,

        ## FilePath will be used. ##AltText?
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [System.String] $Text = $Path,

        [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
        [Alias('PixelHeight')]
        [System.UInt32] $Height,

        [Parameter(ValueFromPipelineByPropertyName, Position = 3)]
        [Alias('PixelWidth')]
        [System.UInt32] $Width,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $AsPercent
    )
    begin {

        if ($AsPercent -and ($Height -le 0 -or $Height -gt 100)) {

            throw ("Height with '-AsPercent' cannot be less-than or equal to 0% and/or greater than 100%.")
        }
        elseif ($AsPercent -and ($Width -le 0 -or $Width -gt 100)) {

            throw ("Width with '-AsPercent' cannot be less-than or equal to 0% and/or greater than 100%.")
        }

        <#! Image.Internal.ps1 !#>

    } #end begin
    process {

        if ($Text.Length -gt 40) {

            $ImageDisplayName = '{0}[..]' -f $Text.Substring(0, 36)
        }
        else {

            $ImageDisplayName = $Text
        }

        WriteLog -Message ($localized.ProcessingImage -f $ImageDisplayName)
        return (New-PScriboImage @PSBoundParameters)

    } #end process
} #end function Image

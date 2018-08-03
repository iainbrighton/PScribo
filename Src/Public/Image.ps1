function Image {
    <#
    .SYNOPSIS
        Initializes a new PScribo Image object.
#>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Image file path
        [Parameter(Mandatory, ParameterSetName = 'Path', ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [System.String] $Path,

        ## Image web uri
        [Parameter(Mandatory, ParameterSetName = 'Uri', ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [System.String] $Uri,

        ## Image width (in pixels)
        [Parameter(Mandatory, ParameterSetName = 'Path', ValueFromPipelineByPropertyName, Position = 1)]
        [Parameter(Mandatory, ParameterSetName = 'Uri', ValueFromPipelineByPropertyName, Position = 1)]
        [System.UInt32] $Height,

        ## Image width (in pixels)
        [Parameter(Mandatory, ParameterSetName = 'Path', ValueFromPipelineByPropertyName, Position = 2)]
        [Parameter(Mandatory, ParameterSetName = 'Uri', ValueFromPipelineByPropertyName, Position = 2)]
        [System.UInt32] $Width,

        ## Image MIME type
        [Parameter(Mandatory, ParameterSetName = 'Uri', ValueFromPipelineByPropertyName)]
        [ValidateSet('bmp','gif','jpeg','tiff','png')]
        [System.String] $MimeType,

        ## Image alignment
        [Parameter(ParameterSetName = 'Path', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'Uri', ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right')]
        [System.String] $Align = 'Left',

        ## Image AltText
        [Parameter(ParameterSetName = 'Path', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'Uri', ValueFromPipelineByPropertyName)]
        [System.String] $Text = $Path,

        [Parameter(ParameterSetName = 'Path', ValueFromPipelineByPropertyName)]
        [Parameter(ParameterSetName = 'Uri', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {

        if ($PSBoundParameters.ContainsKey('Uri')) {

            if (-not $PSBoundParameters.ContainsKey('Text')) {

                $Text = $Uri;
            }
        }

        $PSBoundParameters['Text'] = $Text;


        <#! Image.Internal.ps1 !#>

    } #end begin
    process {

        $imageDisplayName = $Text;
        if ($Text.Length -gt 40) {

            $imageDisplayName = '{0}[..]' -f $Text.Substring(0, 36);
        }

        WriteLog -Message ($localized.ProcessingImage -f $ImageDisplayName);
        return (New-PScriboImage @PSBoundParameters);

    } #end process
} #end function Image

function Image {
<#
    .SYNOPSIS
        Initializes a new PScribo Image object.
#>
    [CmdletBinding(DefaultParameterSetName = 'PathSize')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## Local file path to the image
        [Parameter(Mandatory, ParameterSetName = 'PathSize')]
        [Parameter(Mandatory, ParameterSetName = 'PathPercent')]
        [System.String] $Path,

        ## Remote image web URL path
        [Parameter(Mandatory, ParameterSetName = 'UriSize')]
        [Parameter(Mandatory, ParameterSetName = 'UriPercent')]
        [System.String] $Uri,

        ## Base64 encoded image file
        [Parameter(Mandatory, ParameterSetName = 'Base64Size')]
        [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
        [System.String] $Base64,

        ## Specifies required the image pixel width
        [Parameter(ParameterSetName = 'PathSize')]
        [Parameter(ParameterSetName = 'UriSize')]
        [Parameter(ParameterSetName = 'Base64Size')]
        [System.UInt32] $Height,

        ## Specifies required the image pixel width
        [Parameter(ParameterSetName = 'PathSize')]
        [Parameter(ParameterSetName = 'UriSize')]
        [Parameter(ParameterSetName = 'Base64Size')]
        [System.UInt32] $Width,

        ## Specifies the required scaling percentage
        [Parameter(Mandatory, ParameterSetName = 'PathPercent')]
        [Parameter(Mandatory, ParameterSetName = 'UriPercent')]
        [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
        [System.UInt32] $Percent,

        ## Image alignment
        [Parameter()]
        [ValidateSet('Left','Center','Right')]
        [System.String] $Align = 'Left',

        ## Accessibility image description
        [Parameter(Mandatory, ParameterSetName = 'Base64Size')]
        [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
        [Parameter(ParameterSetName = 'PathSize')]
        [Parameter(ParameterSetName = 'PathPercent')]
        [Parameter(ParameterSetName = 'UriSize')]
        [Parameter(ParameterSetName = 'UriPercent')]
        [System.String] $Text,

        ## Internal image Id
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    process
    {
        if ($PSBoundParameters.ContainsKey('Path'))
        {
            $Uri = Resolve-ImageUri -Path $Path
            $null = $PSBoundParameters.Remove('Path')
            $PSBoundParameters['Uri'] = $Uri
        }
        elseif ($PSBoundParameters.ContainsKey('Uri'))
        {
            $Uri = Resolve-ImageUri -Path $Uri
        }
        elseif ($PSBoundParameters.ContainsKey('Base64'))
        {
            $Uri = Resolve-ImageUri -Path 'about:blank'
        }

        if (-not ($PSBoundParameters.ContainsKey('Text')))
        {
            $Text = $Uri
        }

        $imageDisplayName = $Text
        if ($Text.Length -gt 40)
        {
            $imageDisplayName = '{0}[..]' -f $Text.Substring(0, 36)
        }

        Write-PScriboMessage -Message ($localized.ProcessingImage -f $ImageDisplayName)
        return (New-PScriboImage @PSBoundParameters)
    }
}

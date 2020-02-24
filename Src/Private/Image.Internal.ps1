        #region Image Private Functions

        function New-PScriboImage {
            <#
            .SYNOPSIS
                Initializes a new PScribo Image object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding(DefaultParameterSetName = 'UriSize')]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param
            (
                [Parameter(Mandatory, ParameterSetName = 'UriSize')]
                [Parameter(Mandatory, ParameterSetName = 'UriPercent')]
                [System.String] $Uri,

                [Parameter(Mandatory, ParameterSetName = 'Base64Size')]
                [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
                [System.String] $Base64,

                [Parameter(ParameterSetName = 'UriSize')]
                [Parameter(ParameterSetName = 'Base64Size')]
                [System.UInt32] $Height,

                [Parameter(ParameterSetName = 'UriSize')]
                [Parameter(ParameterSetName = 'Base64Size')]
                [System.UInt32] $Width,

                [Parameter(Mandatory, ParameterSetName = 'UriPercent')]
                [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
                [System.UInt32] $Percent,

                [Parameter()]
                [ValidateSet('Left','Center','Right')]
                [System.String] $Align = 'Left',

                [Parameter(Mandatory, ParameterSetName = 'Base64Size')]
                [Parameter(Mandatory, ParameterSetName = 'Base64Percent')]
                [Parameter(ParameterSetName = 'UriSize')]
                [Parameter(ParameterSetName = 'UriPercent')]
                [System.String] $Text = $Uri,

                [Parameter()]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = [System.Guid]::NewGuid().ToString()
            )
            process
            {
                $imageNumber = [System.Int32] $pscriboDocument.Properties['Images']++;

                if ($PSBoundParameters.ContainsKey('Uri'))
                {
                    $imageBytes = GetImageUriBytes -Uri $Uri
                }
                elseif ($PSBoundParameters.ContainsKey('Base64'))
                {
                    $imageBytes = [System.Convert]::FromBase64String($Base64)
                }

                $image = GetImageFromBytes -Bytes $imageBytes

                if ($PSBoundParameters.ContainsKey('Percent'))
                {
                    $Width = ($image.Width / 100) * $Percent
                    $Height = ($image.Height / 100) * $Percent
                }
                elseif (-not ($PSBoundParameters.ContainsKey('Width')) -and (-not $PSBoundParameters.ContainsKey('Height')))
                {
                    $Width = $image.Width
                    $Height = $image.Height
                }

                $pscriboImage = [PSCustomObject] @{
                    Id          = $Id;
                    ImageNumber = $imageNumber;
                    Text        = $Text
                    Type        = 'PScribo.Image';
                    Bytes       = $imageBytes;
                    Uri         = $Uri;
                    Name        = 'Img{0}' -f $imageNumber;
                    Align       = $Align;
                    MIMEType    = GetImageMimeType -Image $image
                    WidthEm     = ConvertPxToEm -Pixel $Width;
                    HeightEm    = ConvertPxToEm -Pixel $Height;
                    Width       = $Width;
                    Height      = $Height;
                }
                return $pscriboImage;
            }
        }

        function ResolveImagePath {
        <#
            .SYNOPSIS
                Converts an image path into a Uri.
            .NOTES
                A Uri includes information about whether the path is local etc. This is useful for plugins
                to be able to determine whether to embed images or not.
        #>
            [CmdletBinding()]
            [OutputType([System.Uri])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
                [System.String] $Path
            )
            process
            {
                if (Test-Path -Path $Path)
                {
                    $Path = Resolve-Path -Path $Path
                }
                $uri = New-Object -TypeName System.Uri -ArgumentList @($Path)
                return $uri
            }
        }

        function GetImageMimeType {
            [CmdletBinding()]
            [OutputType([System.String])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Drawing.Image] $Image
            )
            process
            {
                if ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Jpeg))
                {
                    return 'image/jpeg'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Png))
                {
                    return 'image/png'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Bmp))
                {
                    return 'image/bmp'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Emf))
                {
                    return 'image/emf'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Gif))
                {
                    return 'image/gif'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Icon))
                {
                    return 'image/icon'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Tiff))
                {
                    return 'image/tiff'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Wmf))
                {
                    return 'image/wmf'
                }
                elseif ($Image.RawFormat.Equals([System.Drawing.Imaging.ImageFormat]::Exif))
                {
                    return 'image/exif'
                }
                else
                {
                    return 'image/unknown'
                }
            }
        }

        function GetImageUriBytes {
        <#
            .SYNOPSIS
                Gets an image's content as a byte[]
        #>
            [CmdletBinding()]
            [OutputType([System.Byte[]])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
                [System.Uri] $Uri
            )
            process
            {
                try
                {
                    $webClient = New-Object -TypeName 'System.Net.WebClient'
                    [System.IO.Stream] $contentStream = $webClient.OpenRead($uri.AbsoluteUri)
                    [System.IO.MemoryStream] $memoryStream = New-Object System.IO.MemoryStream
                    $contentStream.CopyTo($memoryStream)
                    return $memoryStream.ToArray()
                }
                catch
                {
                    $_
                }
                finally
                {
                    if ($null -ne $memoryStream) { $memoryStream.Close() }
                    if ($null -ne $contentStream) { $contentStream.Close() }
                    if ($null -ne $webClient) { $webClient.Dispose() }
                }
            }
        }

        function GetImageFromBytes {
        <#
            .SYNOPSIS
                Creates an image from a byte[]
        #>
            [CmdletBinding()]
            [OutputType([System.Drawing.Image])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Byte[]] $Bytes
            )
            process
            {
                try
                {
                    [System.IO.MemoryStream] $memoryStream = New-Object -TypeName 'System.IO.MemoryStream' -ArgumentList @(,$Bytes)
                    [System.Drawing.Image] $image = [System.Drawing.Image]::FromStream($memoryStream)
                    Write-Output -InputObject $image
                }
                catch
                {
                    $_
                }
                finally
                {
                    if ($null -ne $memoryStream) { $memoryStream.Close() }
                }
            }
        }


        function GetPScriboImage {
        <#
            .SYNOPSIS
                Retrieves PScribo.Images in a document/section
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSObject])]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
                [System.Management.Automation.PSObject[]] $Section,

                [Parameter(ValueFromPipelineByPropertyName)]
                [System.String[]] $Id
            )
            process
            {
                foreach ($subSection in $Section)
                {
                    if ($subSection.Type -eq 'PScribo.Image')
                    {
                        if ($PSBoundParameters.ContainsKey('Id'))
                        {
                            if ($subSection.Id -in $Id)
                            {
                                Write-Output -InputObject $subSection
                            }
                        }
                        else
                        {
                            Write-Output -InputObject $subSection
                        }
                    }
                    elseif ($subSection.Type -eq 'PScribo.Section')
                    {
                        ## Recursively search subsections
                        $PSBoundParameters['Section'] = $subSection.Sections
                        GetPScriboImage @PSBoundParameters
                    }
                }
            }
        }

        #endregion image Private Functions

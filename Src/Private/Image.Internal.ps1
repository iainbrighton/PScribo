#region Image Private Functions

function New-PScriboImage {
    <#
    .SYNOPSIS
        Initializes a new PScribo Image object.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    #[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
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

        if (-not ([System.String]::IsNullOrEmpty($Text))) {
            $Text = $Path.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper()
        }

    } #end begin
    process {

        $imageNumber = [System.Int32] $pscriboDocument.Properties['Images']++;

        if ($PSBoundParameters.ContainsKey('Path')) {

            $imageUri = ResolveImagePath -Path $Path;
        }
        else {

            $imageUri = ResolveImagePath -Path $Uri;
        }

        if (-not $PSBoundParameters.ContainsKey('MimeType')) {

            $MimeType = GetImageMimeType -Path $Path;
        }

        $pscriboImage = [PSCustomObject] @{
            Id          = $Id;
            ImageNumber = $imageNumber;
            Text        = $Text
            Type        = 'PScribo.Image';
            Uri         = $imageUri;
            Name        = 'Img{0}' -f $imageNumber;
            ##RefID       = 'Img{0}' -f $imageNumber
            MIMEType    = $MimeType;
            ## Bytes       = [System.IO.File]::ReadAllBytes($imageUri.LocalPath);

            WidthEm     = ConvertPxToEm -Pixel $Width;
            HeightEm    = ConvertPxToEm -Pixel $Height;
            Width       = $Width;
            Height      = $Height;
        }

        return $pscriboImage;

    } #end process
} #end function New-PScriboImage


function GetImageMimeType {
    <#
        .SYNOPSIS
            Returns an image's MIME type based upon the file extension.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.String] $Path
    )
    process {

        $fileInfo = Get-Item -Path $Path;
        $mimeTypes = @{
            '.bmp'  = 'bmp';
            '.cod'  = 'cis-cod';
            '.gif'  = 'gif';
            '.ief'  = 'ief';
            '.jpe'  = 'jpeg';
            '.jpeg' = 'jpeg';
            '.jpg'  = 'jpeg';
            '.jfif' = 'pipeg';
            '.svg'  = 'svg+sml';
            '.tif'  = 'tiff';
            '.tiff' = 'tiff';
            '.ras'  = 'x-cmu-raster';
            '.cmx'  = 'x-cmx';
            '.ico'  = 'x-icon';
            '.png'  = 'png';
            '.pnm'  = 'x-portable-anymap';
            '.pbm'  = 'x-portable-bitmap';
            '.pgm'  = 'x-portable-graymap';
            '.ppm'  = 'x-portable-pixmap';
            '.rgb'  = 'x-rgb';
            '.xbm'  = 'x-xbitmap';
            '.xpm'  = 'x-xbitmap';
            '.xwd'  = 'x-xwindowdump';
        }
        return $mimeTypes[$fileInfo.Extension];

    } #end process
} #end function

function GetPScriboImage {
<#
    .SYNOPSIS
        Retrieves PScribo.Images in a document/section
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSObject])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.PSObject] $Section,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $Id
    )
    process {

        if ($PSBoundParameters.ContainsKey('Id')) {

            $Section.Sections |
                Where-Object { ($_.Type -eq 'PScribo.Image') -and ($_.Id -in $Id) } |
                    Write-Output;
        }
        else {

            $Section.Sections |
                Where-Object { $_.Type -eq 'PScribo.Image' } |
                    Write-Output;
        }

        ## Recursively search subsections
        $Section.Sections |
            Where-Object { $_.Type -eq 'PScribo.Section' } |
                ForEach-Object {

                    $PSBoundParameters['Section'] = $_;
                    GetPScriboImage @PSBoundParameters;
                }

    }
} #end function GetPScriboImage

#endregion image Private Functions

function GetImageUriBytes {
<#
    .SYNOPSIS
        Gets a web image's content as a byte[]
#>
    [CmdletBinding()]
    [OutputType([System.Byte[]])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Uri] $Uri
    )
    process {

        try {
            $webClient = New-Object -TypeName System.Net.WebClient;
            [System.IO.Stream] $contentStream = $webClient.OpenRead($uri.AbsoluteUri);
            [System.IO.MemoryStream] $memoryStream = New-Object System.IO.MemoryStream;
            $contentStream.CopyTo($memoryStream);
            return $memoryStream.ToArray();
        }
        catch {
            $_
        }
        finally {
            $memoryStream.Close();
            $contentStream.Close();
            $webClient.Dispose();
        }

    } #end process
} #end function

function WriteImageBytes {
<#
    .SYNOPSIS
        Writes an image (byte[]) to file
#>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Byte[]] $Bytes,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Path
    )
    begin {

        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path);

    }
    process {

        try {
            [System.IO.FileStream] $fileStream = New-Object System.IO.FileStream @($Path, [System.IO.FileMode]::Create);
            $fileStream.WriteByte($Bytes, 0, $Bytes.Length);
        }
        catch {
            $_
        }
        finally {
            $fileStream.Close();
        }

    } #end process
} #end function

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
        param (
            [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
            [System.String] $Path
        )
        process {

            if (Test-Path -Path $Path) {

                $Path = Resolve-Path -Path $Path;
            }

            $uri = New-Object -TypeName System.Uri @($Path);

            return $uri;

        }
    } #end function

#endregion image Private Functions

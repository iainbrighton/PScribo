#region Image Private Functions

function New-PScriboImage {
    <#
    .SYNOPSIS
        Initializes a new PScribo Image object.
    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
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
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
    )
    begin {

        if (-not ([System.String]::IsNullOrEmpty($Text))) {
            $Text = $Path.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper()
        }

    } #end begin
    process {

        $imageNumber = [System.Int32] $pscriboDocument.Properties['Image']++
        $typeName = 'PScribo.Image'
        $refID = ('Img{0}' -f $imageNumber)
        $fileItem = Get-Item -Path $Path

        $imageDetail = Get-ImageSize -FilePath $fileItem.FullName
        if ($Height) {

            $imageDetail.PixelHeight = $Height
        }
        if ($Width) {
            $imageDetail.Pixelwidth = $Width
        }

        $pscriboImage = [PSCustomObject] @{
            Id          = $Id
            ImageNumber = $imageNumber
            Text        = $Text
            Type        = $typeName
            Path        = $fileItem.FullName
            Name        = '{0}{1}' -f $RefID, $fileItem.Extension
            RefID       = $refID
            MIMEType    = Get-MimeType -FileInfo $fileItem
            WidthEm     = ConvertPxToEm -Pixel $imageDetail.PixelWidth
            HeightEm    = ConvertPxToEm -Pixel $imageDetail.PixelHeight
            Width       = $imageDetail.PixelWidth
            Height      = $imageDetail.PixelHeight
        }

        return $pscriboImage

    } #end process
} #end function New-PScriboImage

function Get-ImageSize {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.String] $FilePath
    )
    if (-not (Test-Path $FilePath)) {

        throw ('No file found at {0}' -f $FilePath)

    }
    else {

        # load the image
        $Image = [System.Drawing.Image]::FromFile($FilePath)
        [int]$iwidth = $Image.width
        [int]$iheight = $Image.height
        $image.Dispose()
        $ImageDetails = @{
            FilePath    = $FilePath
            PixelWidth  = $iwidth
            PixelHeight = $iheight
        }
        return $ImageDetails
    }
} #end function

function Get-MimeType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [System.IO.FileInfo] $FileInfo
    )
    begin {

        Add-Type -AssemblyName 'System.Web'
        [System.String] $mimeType = $null
    }
    process {

         if ($FileInfo.Exists) {
#             ## Requires PowerShell v4.0 (.NET Framework 4.5 dependency)
             $mimeType = [System.Web.MimeMapping]::GetMimeMapping($FileInfo.FullName)
         }
         else {
             $mimeType = 'false'
         }
    }
    end {

        return $mimeType
    }
}


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
                    Write-Output
        }
        else {

            $Section.Sections |
                Where-Object { $_.Type -eq 'PScribo.Image' } |
                    Write-Output
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

function Get-ImageMimeType
{
<#
    .SYNOPSIS
        Returns an image's Mime type
#>
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

#region Image Private Functions

function New-PScriboImage 
{
  <#
      .SYNOPSIS
      Initializes a new PScribo Image object.
      .NOTES
      This is an internal function and should not be called directly.
  #>
  [CmdletBinding()]
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
  [OutputType([System.Management.Automation.PSCustomObject])]
  param (
    ## FilePath
    [Parameter(Mandatory,ValueFromPipelineByPropertyName, Position = 0)]
    [System.String] $FilePath,
    ## FilePath will be used.
    [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
    [AllowNull()]
    [System.String] $Text = $null,
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = [System.Guid]::NewGuid().ToString()
  )
  begin {

    if (-not ([string]::IsNullOrEmpty($Text))) 
    {
      $Text = $FilePath.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper()
    }

  } #end begin
  process {
    $ImageNumber=0
    IF ($pscriboDocument.Properties['Image'] -gt 0){$ImageNumber=$pscriboDocument.Properties['Image']}
    $typeName = 'PScribo.Image'
    $RefID=('Img{0}' -f $ImageNumber)
    $ImageDetails=ImageSizeInEMU -FilePath $FilePath
    $pscriboDocument.Properties['Image']++
    $pscriboImage = [PSCustomObject] @{
      ID       = $Id
      ImageNumber=$ImageNumber
      Text     = $Text
      Type     = $typeName
      FilePath = $FilePath
      RefID= $RefId
      MIME = Get-MimeType -CheckFile $FilePath
      Name = ('{0}{1}' -f $RefID, $(Get-Item -Path $FilePath).Extension)
      EMUWidth=$imageDetails.EMUWidth
      EMUHeight=$imageDetails.EMUHeight
      PixelWidth=$imageDetails.PixelWidth
      PixelHeight=$ImageDetails.PixelHeight
    }

    return $pscriboImage

  } #end process
} #end function New-PScriboImage
Function ImageSizeInEMU 
{
  param(
    [Parameter(Mandatory)]$FilePath
  )
  IF(-not (Test-Path $FilePath))
  {
    Throw ('No file found at {0}' -f $FilePath)
  }
  Else
  {
    $drawfile = Get-ChildItem -Path "$((Get-ChildItem -Path Env:\windir).Value)\assembly" -Filter *drawing.dll -Recurse
    $dllpath = (Get-Command $($drawfile.Fullname)).definition
    $null = [Reflection.Assembly]::LoadFrom($dllpath)
    # load the image
    $Image = [Drawing.Image]::FromFile($FilePath)
    [int]$iwidth = $Image.width
    [int]$iheight = $Image.height
    #convert pixels to EMU
    $ImageDetails = @{
      FilePath    = $FilePath
      PixelWidth  = $iwidth
      PixelHeight = $iheight
      EMUWidth    = [int][Math]::Round(($iwidth * 9525),0)
      EMUHeight   = [int][Math]::Round(($iheight * 9525),0)
    }
    return $ImageDetails
  }
}
 
function Get-MimeType() 
{ 
  param([parameter(Mandatory = $true, ValueFromPipeline = $true)][ValidateNotNullorEmpty()][System.IO.FileInfo]$CheckFile) 
  begin { 
    Add-Type -AssemblyName 'System.Web'         
    [System.IO.FileInfo]$check_file = $CheckFile 
    [string]$mime_type = $null 
  } 
  process { 
    if ($check_file.Exists) 
    {
      $mime_type = [System.Web.MimeMapping]::GetMimeMapping($check_file.FullName)
    } 
    else 
    {
      $mime_type = 'false'
    } 
  } 
  end { return $mime_type } 
}
#endregion image Private Functions
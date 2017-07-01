function OutWord 
{
  <#
      .SYNOPSIS
      Microsoft Word output plugin for PScribo.
      .DESCRIPTION
      Outputs a Word document representation of a PScribo document object.
  #>
  [CmdletBinding()]
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
  [OutputType([System.IO.FileInfo])]
  param (
    ## ThePScribo document object to convert to a text document
    [Parameter(Mandatory, ValueFromPipeline)]
    [System.Object] $Document,

    ## Output directory path for the .txt file
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
    [ValidateNotNull()]
    [System.String] $Path,

    ### Hashtable of all plugin supported options
    [Parameter()]
    [AllowNull()]
    [System.Collections.Hashtable] $Options
  )
  begin {

    $pluginName = 'Word'

    <#! OutWord.Internal.ps1 !#>

  }
  process {

    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $Images=@()
    WriteLog -Message ($localized.DocumentProcessingStarted -f $Document.Name)
    $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $xmlnswpdrawing = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    $xmlnsdrawing = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    $xmlnspicture = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    $xmlnsrelationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $xmlnsofficeword14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
    $xmlnsmath = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    $xmlDocument = New-Object -TypeName 'System.Xml.XmlDocument'
    [ref] $null = $xmlDocument.AppendChild($xmlDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
    $documentXml = $xmlDocument.AppendChild($xmlDocument.CreateElement('w', 'document', $xmlnsMain))
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:xml', 'http://www.w3.org/XML/1998/namespace')
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:pic', $xmlnspicture)
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:wp', $xmlnswpdrawing)
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:a', $xmlnsdrawing)
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:r', $xmlnsrelationships)
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:m', $xmlnsmath)
    [ref] $null = $xmlDocument.DocumentElement.SetAttribute('xmlns:a14', $xmlnsofficeword14)
        
    $body = $documentXml.AppendChild($xmlDocument.CreateElement('w', 'body', $xmlnsMain))
    ## Setup the document page size/margins
    $sectionPrParams = @{
      PageHeight       = $Document.Options['PageHeight']
      PageWidth        = $Document.Options['PageWidth']
      PageMarginTop    = $Document.Options['MarginTop']
      PageMarginBottom = $Document.Options['MarginBottom']
      PageMarginLeft   = $Document.Options['MarginLeft']
      PageMarginRight  = $Document.Options['MarginRight']
    }
    [ref] $null = $body.AppendChild((GetWordSectionPr @sectionPrParams -XmlDocument $xmlDocument))
    foreach ($s in $Document.Sections.GetEnumerator()) 
    {
      if ($s.Id.Length -gt 40) 
      {
        $sectionId = '{0}[..]' -f $s.Id.Substring(0,36)
      }
      else 
      {
        $sectionId = $s.Id
      }
      $currentIndentationLevel = 1
      if ($null -ne $s.PSObject.Properties['Level']) 
      {
        $currentIndentationLevel = $s.Level +1
      }
      WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel
      switch ($s.Type) {
        'PScribo.Section' 
        {
          $s | OutWordSection -RootElement $body -XmlDocument $xmlDocument
        }
        'PScribo.Paragraph' 
        {
          [ref] $null = $body.AppendChild((OutWordParagraph -Paragraph $s -XmlDocument $xmlDocument))
        }
        'PScribo.Image' 
        { 
          $Images+=@($s)
          [ref] $null = $body.AppendChild((OutWordImage -Image $s -XmlDocument $xmlDocument))
        }
        'PScribo.PageBreak' 
        {
          [ref] $null = $body.AppendChild((OutWordPageBreak -PageBreak $s -XmlDocument $xmlDocument))
        }
        'PScribo.LineBreak' 
        {
          [ref] $null = $body.AppendChild((OutWordLineBreak -LineBreak $s -XmlDocument $xmlDocument))
        }
        'PScribo.Table' 
        {
          OutWordTable -Table $s -XmlDocument $xmlDocument -Element $body
        }
        'PScribo.TOC' 
        {
          [ref] $null = $body.AppendChild((OutWordTOC -TOC $s -XmlDocument $xmlDocument))
        }
        'PScribo.BlankLine' 
        {
          OutWordBlankLine -BlankLine $s -XmlDocument $xmlDocument -Element $body
        }
        Default 
        {
          WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning
        }
      } #end switch
    } #end foreach
    ## Generate the Word 'styles.xml' document part
    $stylesXml = OutWordStylesDocument -Styles $Document.Styles -TableStyles $Document.TableStyles
    ## Generate the Word 'settings.xml' document part
    if (($Document.Properties['TOCs']) -and ($Document.Properties['TOCs'] -gt 0)) 
    {
      ## We have a TOC so flag to update the document when opened
      $settingsXml = OutWordSettingsDocument -UpdateFields
    }
    else 
    {
      $settingsXml = OutWordSettingsDocument
    }
    #Convert relative or PSDrive based path to the absolute filesystem path
    $AbsolutePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $destinationPath = Join-Path -Path $AbsolutePath -ChildPath ('{0}.docx' -f $Document.Name)
    if ($PSVersionTable.PSEdition -ne 'Core') 
    {
      ## WindowsBase.dll is not included in Core PowerShell
      Add-Type -AssemblyName WindowsBase
    }
    try 
    {
      $package = [System.IO.Packaging.Package]::Open($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)
    }
    catch 
    {
      WriteLog -Message ($localized.OpenPackageError -f $destinationPath) -IsWarning
      throw $_
    }
    ## Create document.xml part
    $documentUri = New-Object -TypeName System.Uri -ArgumentList ('/word/document.xml', [System.UriKind]::Relative)
    WriteLog -Message ($localized.ProcessingDocumentPart -f $documentUri)
    $documentPart = $package.CreatePart($documentUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($documentPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
    $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
    WriteLog -Message ($localized.WritingDocumentPart -f $documentUri)
    $xmlDocument.Save($xmlWriter)
    $xmlWriter.Dispose()
    $streamWriter.Close()

    ## Create styles.xml part
    $stylesUri = New-Object -TypeName System.Uri -ArgumentList ('/word/styles.xml', [System.UriKind]::Relative)
    WriteLog -Message ($localized.ProcessingDocumentPart -f $stylesUri)
    $stylesPart = $package.CreatePart($stylesUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml')
    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($stylesPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
    $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
    WriteLog -Message ($localized.WritingDocumentPart -f $stylesUri)
    $stylesXml.Save($xmlWriter)
    $xmlWriter.Dispose()
    $streamWriter.Close()

    ## Create settings.xml part
    $settingsUri = New-Object -TypeName System.Uri -ArgumentList ('/word/settings.xml', [System.UriKind]::Relative)
    WriteLog -Message ($localized.ProcessingDocumentPart -f $settingsUri)
    $settingsPart = $package.CreatePart($settingsUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
    $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($settingsPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
    $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
    WriteLog -Message ($localized.WritingDocumentPart -f $settingsUri)
    $settingsXml.Save($xmlWriter)
    $xmlWriter.Dispose()
    $streamWriter.Close()
                

        
    ## Create the Package relationships
    WriteLog -Message $localized.GeneratingPackageRelationships
    [ref] $null = $package.CreateRelationship($documentUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'rId1')
    [ref] $null = $documentPart.CreateRelationship($stylesUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'rId1')
    [ref] $null = $documentPart.CreateRelationship($settingsUri, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'rId2')
    ## Process images
    $i=0
    While ($images.Count -gt $i)
    {      
      $CurrentImage=$Images[$i]
      $URI = ('/word/media/{0}' -f $CurrentImage.Name)
      $partName = New-Object -TypeName System.Uri -ArgumentList ($URI, [System.UriKind]'Relative')
      $part = $package.CreatePart($partName, $CurrentImage.MIME)
      $bytes = [System.IO.File]::ReadAllBytes($CurrentImage.FilePath)
      $stream = $part.GetStream()
      $stream.Write($bytes, 0, $bytes.Length)
      $stream.Close()
      [ref] $null = $documentPart.CreateRelationship($partName, [System.IO.Packaging.TargetMode]::Internal, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', $CurrentImage.RefId)
      $i++
    }

		
    $package.Flush()
    $package.Close()

    $stopwatch.Stop()
    WriteLog -Message ($localized.DocumentProcessingCompleted -f $Document.Name)
    WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds)
    ## Return the file reference to the pipeline
    Write-Output -InputObject (Get-Item -Path $destinationPath)

  } #end process
} #end function OutWord

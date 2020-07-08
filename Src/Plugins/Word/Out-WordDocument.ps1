function Out-WordDocument
{
<#
    .SYNOPSIS
        Microsoft Word output plugin for PScribo.

    .DESCRIPTION
        Outputs a Word document representation of a PScribo document object.
  #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'pluginName')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','Options')]
    [OutputType([System.IO.FileInfo])]
    param
    (
        ## ThePScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ## Output directory path for the .txt file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNull()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter()]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    process
    {
        $pluginName = 'Word'
        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Document.Name)

        ## Generate the Word 'document.xml' document part
        $documentXml = Get-WordDocument -Document $Document

        ## Generate the Word 'styles.xml' document part
        $stylesXml = Get-WordStylesDocument -Styles $Document.Styles -TableStyles $Document.TableStyles

        ## Generate the Word 'settings.xml' document part
        $updateFields = (($Document.Properties['TOCs']) -and ($Document.Properties['TOCs'] -gt 0))
        $settingsXml = Get-WordSettingsDocument -UpdateFields:$updateFields

        #Convert relative or PSDrive based path to the absolute filesystem path
        $absolutePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $destinationPath = Join-Path -Path $absolutePath ('{0}.docx' -f $Document.Name)
        if ($PSVersionTable['PSEdition'] -ne 'Core')
        {
            ## "Unable to determine the identity of the domain" fix (https://github.com/AsBuiltReport/AsBuiltReport.Core/issues/17)
            $currentAppDomain = [System.Threading.Thread]::GetDomain()
            $securityIdentityField = $currentAppDomain.GetType().GetField("_SecurityIdentity", ([System.Reflection.BindingFlags]::Instance -bOr [System.Reflection.BindingFlags]::NonPublic))
            $replacementEvidence = New-Object -TypeName 'System.Security.Policy.Evidence'
            $securityZoneMyComputer = New-Object -TypeName 'System.Security.Policy.Zone' -ArgumentList @([System.Security.SecurityZone]::MyComputer)
            $replacementEvidence.AddHost($securityZoneMyComputer)
            $securityIdentityField.SetValue($currentAppDomain, $replacementEvidence)
            ## WindowsBase.dll is not included in Core PowerShell
            Add-Type -AssemblyName WindowsBase
        }
        try
        {
            $package = [System.IO.Packaging.Package]::Open($destinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite)
        }
        catch
        {
            Write-PScriboMessage -Message ($localized.OpenPackageError -f $destinationPath) -IsWarning
            throw $_
        }

        ## Create document.xml part
        $documentUri = New-Object -TypeName System.Uri -ArgumentList ('/word/document.xml', [System.UriKind]::Relative)
        Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $documentUri)
        $documentPart = $package.CreatePart($documentUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
        $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($documentPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
        Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $documentUri)
        $documentXml.Save($xmlWriter)
        $xmlWriter.Dispose()
        $streamWriter.Close()

        ## Create styles.xml part
        $stylesUri = New-Object -TypeName System.Uri -ArgumentList ('/word/styles.xml', [System.UriKind]::Relative)
        Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $stylesUri)
        $stylesPart = $package.CreatePart($stylesUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml')
        $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($stylesPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
        Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $stylesUri)
        $stylesXml.Save($xmlWriter)
        $xmlWriter.Dispose()
        $streamWriter.Close()

        ## Create settings.xml part
        $settingsUri = New-Object -TypeName System.Uri -ArgumentList ('/word/settings.xml', [System.UriKind]::Relative)
        Write-PScriboMessage -Message ($localized.ProcessingDocumentPart -f $settingsUri)
        $settingsPart = $package.CreatePart($settingsUri, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')
        $streamWriter = New-Object -TypeName System.IO.StreamWriter -ArgumentList ($settingsPart.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite))
        $xmlWriter = [System.Xml.XmlWriter]::Create($streamWriter)
        Write-PScriboMessage -Message ($localized.WritingDocumentPart -f $settingsUri)
        $settingsXml.Save($xmlWriter)
        $xmlWriter.Dispose()
        $streamWriter.Close()

        Out-WordHeaderFooterDocument -Document $Document -Package $package

        ## Create the Package relationships
        Write-PScriboMessage -Message $localized.GeneratingPackageRelationships
        $officeDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
        $stylesDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
        $settingsDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'

        [ref] $null = $package.CreateRelationship($documentUri, [System.IO.Packaging.TargetMode]::Internal, $officeDocumentUri, 'rId1')
        [ref] $null = $documentPart.CreateRelationship($stylesUri, [System.IO.Packaging.TargetMode]::Internal, $stylesDocumentUri, 'rId1')
        [ref] $null = $documentPart.CreateRelationship($settingsUri, [System.IO.Packaging.TargetMode]::Internal, $settingsDocumentUri, 'rId2')

        $headerDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
        if ($Document.Header.HasFirstPageHeader)
        {
            $firstPageHeaderUri = New-Object -TypeName System.Uri -ArgumentList ('/word/firstPageHeader.xml', [System.UriKind]::Relative)
            [ref] $null = $documentPart.CreateRelationship($firstPageHeaderUri, [System.IO.Packaging.TargetMode]::Internal, $headerDocumentUri, 'rId3')
        }
        if ($Document.Header.HasDefaultHeader)
        {
            $defaultHeaderUri = New-Object -TypeName System.Uri -ArgumentList ('/word/defaultHeader.xml', [System.UriKind]::Relative)
            [ref] $null = $documentPart.CreateRelationship($defaultHeaderUri, [System.IO.Packaging.TargetMode]::Internal, $headerDocumentUri, 'rId4')
        }

        $footerDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer'
        if ($Document.Footer.HasFirstPageFooter)
        {
            $firstPageFooterUri = New-Object -TypeName System.Uri -ArgumentList ('/word/firstPageFooter.xml', [System.UriKind]::Relative)
            [ref] $null = $documentPart.CreateRelationship($firstPageFooterUri, [System.IO.Packaging.TargetMode]::Internal, $footerDocumentUri, 'rId5')
        }
        if ($Document.Footer.HasDefaultFooter)
        {
            $defaultFooterUri = New-Object -TypeName System.Uri -ArgumentList ('/word/defaultFooter.xml', [System.UriKind]::Relative)
            [ref] $null = $documentPart.CreateRelationship($defaultFooterUri, [System.IO.Packaging.TargetMode]::Internal, $footerDocumentUri, 'rId6')
        }

        ## Process images (assuming we have a section, e.g. example03.ps1)
        if ($Document.Sections.Count -gt 0)
        {
            $imageDocumentUri = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            foreach ($image in (Get-PScriboImage -Section $Document.Sections))
            {
                try
                {
                    $uri = ('/word/media/{0}' -f $image.Name)
                    $partName = New-Object -TypeName 'System.Uri' -ArgumentList ($uri, [System.UriKind]'Relative')
                    $part = $package.CreatePart($partName, $image.MimeType)
                    $stream = $part.GetStream()
                    $stream.Write($image.Bytes, 0, $image.Bytes.Length)
                    $stream.Close()
                    [ref] $null = $documentPart.CreateRelationship($partName, [System.IO.Packaging.TargetMode]::Internal, $imageDocumentUri, $image.Name)
                }
                catch
                {
                    throw $_
                }
                finally
                {
                    if ($null -ne $stream) { $stream.Close() }
                }
            }
        }

        $package.Flush()
        $package.Close()
        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)

        if ($stopwatch.Elapsed.TotalSeconds -gt 90)
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeMinutes -f $stopwatch.Elapsed.TotalMinutes)
        }
        else
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeSeconds -f $stopwatch.Elapsed.TotalSeconds)
        }

        Write-Output -InputObject (Get-Item -Path $destinationPath)
    }
}

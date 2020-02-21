#requires -Version 5;
#requires -Modules VirtualEngine.Build;

$psake.use_exit_on_error = $true;

Properties {
    $moduleName = (Get-Item $PSScriptRoot\*.psd1)[0].BaseName;
    $basePath = $psake.build_script_dir;
    $buildDir = 'Release';
    $buildPath = (Join-Path -Path $basePath -ChildPath $buildDir);
    $releasePath = (Join-Path -Path $buildPath -ChildPath $moduleName);
    $thumbprint = '177FC8E667D4C022C7CD9CFDFEB66991890F4090';
    $timeStampServer = 'http://timestamp.digicert.com';
    $exclude = @(
        '.git*',
        '.vscode',
        'Release',
        'Tests',
        'Build.PSake.ps1',
        '*.png',
        '*.md',
        '*.enc',
        'TestResults.xml',
        'appveyor.yml',
        'appveyor-tools'
        'PScribo Test Doc.*',
        'PScriboExample.*',
        'TestResults.xml',
        'System.IO.Packaging.dll'
    );
    $signExclude = @('Examples','en-US');
}

#region functions

function New-PscriboBundle {
    [CmdletBinding()]
    param (
        ## Files to bundle.
        [Parameter(Mandatory)]
        [System.String[]] $Path,

        ## Output filename.
        [Parameter(Mandatory)]
        [System.String] $DestinationPath,

        ## Excluded files.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Exclude = @('*.Tests.ps1','*.Internal.ps1')
    )
    Write-Host ('  Creating bundle header.') -ForegroundColor Cyan;
    Set-Content -Path $DestinationPath -Value "#region PScribo Bundle v$version";
    Add-Content -Path $DestinationPath -Value "#requires -Version 3`r`n";

    ## Import LICENSE
    Write-Host ('  Creating bundle license.') -ForegroundColor Cyan;
    Get-Content -Path "$currentDir\LICENSE" | Add-Content -Path $DestinationPath;

    ## TODO: Support localised bundles, eg en-US and fr-FR
    Write-Host ('  Creating bundle resources.') -ForegroundColor Cyan;
    Add-Content -Path $DestinationPath -Value "`r`n`$localized = DATA {";
    Get-Content -Path "$currentDir\en-US\PScribo.Resources.psd1" | Add-Content -Path $DestinationPath;
    Add-Content -Path $DestinationPath -Value "}`r`n";

    Join-PscriboBundleFile -Path $Path -DestinationPath $DestinationPath -Exclude $Exclude;
    Write-Host ('  Creating bundle footer.') -ForegroundColor Cyan;
    Add-Content -Path $DestinationPath -Value "#endregion PScribo Bundle v$version";
}

function Join-PscriboBundleFile {
    [CmdletBinding()]
    param (
        ## Files to bundle.
        [Parameter(Mandatory)]
        [System.String[]] $Path,

        ## Output filename.
        [Parameter(Mandatory)]
        [System.String] $DestinationPath,

        ## Excluded files.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Exclude = @('*.Tests.ps1')
    )

    process {
        foreach ($file in (Get-ChildItem -Path $Path -Exclude $Exclude)) {
            Write-Host ('   Bundling file "{0}".' -f (TrimPath -Path $file)) -ForegroundColor DarkCyan;
            $internalFunctionContent = "";
            $content = Get-Content -Path $file.FullName -Raw;
            if ($content -match '(?<=<#!\s?)\S+.Internal.ps1(?=\s?!#>)') {

                $internalFunctionDirectory = Join-Path -Path (Split-Path -Path $file.DirectoryName -Parent) -ChildPath Private;
                $internalFunctionPath = Join-Path -Path $internalFunctionDirectory -ChildPath $Matches[0];
                if (Test-Path -Path $internalFunctionPath) {

                    if (Test-Path -Path $internalFunctionPath) {
                        Write-Host ('    Bundling internal file "{0}".' -f (TrimPath -Path $internalFunctionPath)) -ForegroundColor DarkCyan;
                        $internalFunctionContent = Get-Content -Path $internalFunctionPath -Raw;
                        if ($content -match '<#!\s?\S+.Internal.ps1\s?!#>') {
                            $replacementString = $Matches[0];
                            Write-Debug ('Replacing text "{0}"...' -f $replacementString);
                            ## Cannot use the -replace method as it replaces occurences of $_ too ?!
                            $content = $content.Replace($replacementString, $internalFunctionContent.Trim());
                        }
                    }
                }
            }

            ## Now remove empty lines
            $compressedContent = New-Object -TypeName System.Text.StringBuilder;
            $content.Split("`r`n") | ForEach-Object {

                if (-not [System.String]::IsNullOrWhitespace($_)) {

                    [ref] $null = $compressedContent.AppendLine($_);
                }
            }
            $compressedContent.ToString() | Add-Content -Path $DestinationPath;
        }
    }
}

#endregion functions

# Synopsis: Initialises build variables
Task Init {

    # Properties are not available in the script scope.
    Set-Variable manifest -Value (Get-ModuleManifest) -Scope Script;
    Set-Variable version -Value $manifest.Version -Scope Script;
    Write-Host (" Building module '{0}'." -f $manifest.Name) -ForegroundColor Yellow;
    Write-Host (" Building version '{0}'." -f $version) -ForegroundColor Yellow;
} #end task Init

# Synopsis: Cleans the release directory
Task Clean -Depends Init {

    Write-Host (' Cleaning release directory "{0}".' -f $buildPath) -ForegroundColor Yellow;
    if (Test-Path -Path $buildPath) {
        Remove-Item -Path $buildPath -Include * -Recurse -Force;
    }
    [ref] $null = New-Item -Path $buildPath -ItemType Directory -Force;
    [ref] $null = New-Item -Path $releasePath -ItemType Directory -Force;
} #end task Clean

# Synopsis: Invokes Pester tests
Task Test -Depends Init {

    $invokePesterParams = @{
        Path = "$basePath\Tests";
        OutputFile = "$basePath\TestResults.xml";
        OutputFormat = 'NUnitXml';
        Strict = $true;
        PassThru = $true;
        Verbose = $false;
    }
    $testResult = Invoke-Pester @invokePesterParams;
    if ($testResult.FailedCount -gt 0) {
        Write-Error ('Failed "{0}" unit tests.' -f $testResult.FailedCount);
    }
}

# Synopsis: Copies release files to the release directory
Task Deploy -Depends Clean {

    Get-ChildItem -Path $basePath -Exclude $exclude | ForEach-Object {
        Write-Host (' Copying {0}' -f $PSItem.FullName) -ForegroundColor Yellow;
        Copy-Item -Path $PSItem -Destination $releasePath -Recurse;
    }
} #end

# Synopsis: Signs files in release directory
Task Sign -Depends Deploy {

    if (-not (Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object Thumbprint -eq $thumbprint)) {
        ## Decrypt and import code signing cert
        .\appveyor-tools\secure-file.exe -decrypt .\VE_Certificate_2021.pfx.enc -secret $env:certificate_secret
        $certificatePassword = ConvertTo-SecureString -String $env:certificate_secret -AsPlainText -Force
        Import-PfxCertificate -FilePath .\VE_Certificate_2021.pfx -CertStoreLocation 'Cert:\CurrentUser\My' -Password $certificatePassword
    }

    Get-ChildItem -Path $releasePath -Exclude $signExclude | ForEach-Object {
        if ($PSItem -is [System.IO.DirectoryInfo]) {
            Get-ChildItem -Path $PSItem.FullName -Include *.ps* -Recurse | ForEach-Object {
                Write-Host (' Signing {0}' -f $PSItem.FullName) -ForegroundColor Yellow -NoNewline;
                $signResult = Set-ScriptSignature -Path $PSItem.FullName -Thumbprint $thumbprint -TimeStampServer $timeStampServer -ErrorAction Stop;
                Write-Host (' {0}.' -f $signResult.Status) -ForegroundColor Green;
            }

        }
        elseif ($PSItem.Name -like '*.ps*') {
            Write-Host (' Signing {0}' -f $PSItem.FullName) -ForegroundColor Yellow -NoNewline;
            $signResult = Set-ScriptSignature -Path $PSItem.FullName -Thumbprint $thumbprint -TimeStampServer $timeStampServer -ErrorAction Stop;
            Write-Host (' {0}.' -f $signResult.Status) -ForegroundColor Green;
        }
    }
}

Task Version -Depends Deploy {

    $nuSpecPath = Join-Path -Path $releasePath -ChildPath "$ModuleName.nuspec"
    $nuspec = [System.Xml.XmlDocument] (Get-Content -Path $nuSpecPath -Raw)
    $nuspec.Package.MetaData.Version = $version.ToString()
    $nuspec.Save($nuSpecPath)
}

# Synopsis: Publishes release module to PSGallery
Task Publish_PSGallery -Depends Version {

    Publish-Module -Path $releasePath -NuGetApiKey "$env:gallery_api_key" -Verbose
} #end task Publish

# Synopsis: Creates release module Nuget package
Task Package -Depends Build {

    $targetNuSpecPath = Join-Path -Path $releasePath -ChildPath "$ModuleName.nuspec"
    NuGet.exe pack "$targetNuSpecPath" -OutputDirectory "$env:TEMP"
}

# Synopsis: Publish release module to Dropbox repository
Task Publish_Dropbox -Depends Package {

    $targetNuPkgPath = Join-Path -Path "$env:TEMP" -ChildPath "$ModuleName.$version.nupkg"
    $destinationPath = "$env:USERPROFILE\Dropbox\PSRepository"
    Copy-Item -Path "$targetNuPkgPath"-Destination $destinationPath -Force -Verbose
}

# Synopsis: Publish test results to AppVeyor
Task AppVeyor {

    Get-ChildItem -Path "$basePath\*Results*.xml" | Foreach-Object {
        $address = 'https://ci.appveyor.com/api/testresults/nunit/{0}' -f $env:APPVEYOR_JOB_ID
        $source = $_.FullName
        Write-Verbose "UPLOADING TEST FILE: $address $source" -Verbose
        (New-Object 'System.Net.WebClient').UploadFile( $address, $source )
    }
}

Task Default -Depends Init, Clean, Test
Task Build -Depends Default, Deploy, Version, Sign;
Task Publish -Depends Build, Package, Publish_PSGallery
Task Local -Depends Build, Package, Publish_Dropbox

#requires -Version 5;
#requires -Modules VirtualEngine.Build;
$psake.use_exit_on_error = $true;

Properties {
    $currentDir = Resolve-Path -Path (Get-Location -PSProvider FileSystem);
    $basePath = $psake.build_script_dir;
    $buildDir = 'Build';
    $releaseDir = 'Release';
    $company = 'Iain Brighton';
    $author = 'Iain Brighton';
    $thumbprint = '3DACD0F2D1E60EB33EC774B9CFC89A4BEE9037AF';
    $timeStampServer = 'http://timestamp.verisign.com/scripts/timestamp.dll';
}

Task Default -Depends Build;
Task Build -Depends Clean, Setup, Test, Deploy;
Task Stage -Depends Build, Version, Bundle, Sign, Zip;
#Task Publish -Depends Stage, Release;


Task Test {
    $testResult = Invoke-Pester -Path $basePath -OutputFile "$buildPath\TestResult.xml" -OutputFormat NUnitXml -Strict -PassThru;
    if ($testResult.FailedCount -gt 0) {
        Write-Error ('Failed "{0}" unit tests.' -f $testResult.FailedCount);
    }
}


Task Clean {
    Write-Host (' Base directory "{0}".' -f $basePath) -ForegroundColor Yellow;
    ## Remove build directory
    $baseBuildPath = Join-Path -Path $psake.build_script_dir -ChildPath $buildDir;
    if (Test-Path -Path $baseBuildPath) {
        Write-Host (' Removing build base directory "{0}".' -f (TrimPath -Path $baseBuildPath)) -ForegroundColor Yellow;
        Remove-Item $baseBuildPath -Recurse -Force -ErrorAction Stop;
    }
}


Task Setup {
    # Properties are not available in the script scope.
    Set-Variable manifest -Value (Get-ModuleManifest) -Scope Script;
    Set-Variable buildPath -Value (Join-Path -Path $psake.build_script_dir -ChildPath "$buildDir\$($manifest.Name)") -Scope Script;
    Set-Variable releasePath -Value (Join-Path -Path $psake.build_script_dir -ChildPath $releaseDir) -Scope Script;
    $newModuleVersion = New-Object -TypeName System.Version -ArgumentList $manifest.Version.Major, $manifest.Version.Minor,$manifest.Version.Build,(Get-GitRevision);
    Set-Variable version -Value ($newModuleVersion.ToString()) -Scope Script;

    Write-Host (' Building module "{0}".' -f $manifest.Name) -ForegroundColor Yellow;
    Write-Host (' Using version number "{0}".' -f $version) -ForegroundColor Yellow;

    ## Create the build directory
    Write-Host (' Creating build directory "{0}".' -f (TrimPath -Path $buildPath)) -ForegroundColor Yellow;
    [Ref] $null = New-Item $buildPath -ItemType Directory -Force -ErrorAction Stop;

    ## Create the release directory
    if (!(Test-Path -Path $releasePath)) {
        Write-Host (' Creating release directory "{0}".' -f (TrimPath -Path $releasePath)) -ForegroundColor Yellow;
        [Ref] $null = New-Item $releasePath -ItemType Directory -Force -ErrorAction Stop;
    }
}


Task Deploy {
    ## Copy release files
    Write-Host (' Copying release files to build directory "{0}".' -f (TrimPath -Path $buildPath)) -ForegroundColor Yellow;
    $excludedFiles = @(
        '*.Tests.ps1',
        'Build.PSake.ps1',
        '.git*',
        '*.png',
        'Build',
        'Release',
        'readme.md',
        'bin',
        'obj',
        '*.sln',
        '*.suo',
        '*.pssproj',
        'PScribo Test Doc.*',
        'PScriboExample.*',
        'TestResult.xml',
        '.vscode',
        'System.IO.Packaging.dll'
    );
    Get-ModuleFile -Exclude $excludedFiles | ForEach-Object {
        $destinationPath = '{0}{1}' -f $buildPath, $PSItem.FullName.Replace($basePath, '');
        Write-Host ('  Copying release file "{0}".' -f (TrimPath -Path $destinationPath)) -ForegroundColor DarkCyan;
        [Ref] $null = New-Item -ItemType File -Path $destinationPath -Force;
        Copy-Item -Path $PSItem.FullName -Destination $destinationPath -Force;
    }
}


Task Version {
    ## Version module manifest prior to build
    $manifestPath = Join-Path $buildPath -ChildPath "$($manifest.Name).psd1";
    Write-Host (' Versioning module manifest "{0}".' -f $manifestPath) -ForegroundColor Yellow;
    Set-ModuleManifestProperty -Path $manifestPath -Version $version -CompanyName $company -Author $author;
    ## Reload module manifest to ensure the version number is picked back up
    Set-Variable manifest -Value (Get-ModuleManifest -Path $manifestPath) -Scope Script -Force;
}


Task Sign {
    Get-ChildItem -Path $buildPath -Include *.ps* -Recurse -File | % {
        Write-Host (' Signing file "{0}":' -f (TrimPath -Path $PSItem.FullName)) -ForegroundColor Yellow -NoNewline;
        $signResult = Set-ScriptSignature -Path $PSItem.FullName -Thumbprint $thumbprint -TimeStampServer $timeStampServer -ErrorAction Stop;
        Write-Host (' {0}.' -f $signResult.Status) -ForegroundColor Green;
    }
}


Task Zip {
    ## Creates the release files in the $releaseDir
    $zipReleaseName = '{0}-v{1}.zip' -f $manifest.Name, $version;
    $zipPath = Join-Path -Path $releasePath -ChildPath $zipReleaseName;
    Write-Host (' Creating zip file "{0}".' -f (TrimPath -Path $zipPath)) -ForegroundColor Yellow;
    ## Zip the parent directory
    $zipSourcePath = Split-Path -Path $buildPath -Parent;
    $zipFile = New-ZipArchive -Path $zipSourcePath -DestinationPath $zipPath;
    Write-Host (' Zip file "{0}" created.' -f (TrimPath -Path $zipFile.Fullname)) -ForegroundColor Yellow;
}

Task Bundle {
    $bundlePath = Join-Path -Path $releasePath -ChildPath "PScribo-v$version-Bundle.ps1";
    Write-Host (' Creating bundle file "{0}".' -f (TrimPath -Path $bundlePath)) -ForegroundColor Yellow;
    $bundleFiles = "$buildPath\Src\Public","$buildPath\Src\Private","$buildPath\Src\Plugins";
    $excludedFiles = '*.Tests.ps1','*.Internal.ps1','System.IO.Packaging.dll';
    Bundle-File -Path $bundleFiles -DestinationPath $bundlePath -Exclude $excludedFiles -Verbose;
}

<#
Task DocumentBundle {
    $bundlePath = Join-Path -Path $buildPath -ChildPath "PScribo-$version-DocumentBundle.ps1";
    $bundleFiles = "$currentDir\LICENSE","$currentDir\Functions";
    Bundle-File -Path $bundleFiles -DestinationPath $bundlePath -Verbose;
}

Task OutputBundle {
    $bundlePath = Join-Path -Path $buildPath -ChildPath "PScribo-$version-OutputBundle.ps1";
    $bundleFiles = "$currentDir\LICENSE","$currentDir\Plugins";
    Bundle-File -Path $bundleFiles -DestinationPath $bundlePath -Verbose;
}
#>

Task Minify { }


function TrimPath {
<#
    .SYNOPSIS
        Trims a directory path to a relative path to avoid wrapping issues
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.String] $Path
    )
    return ($Path.Replace($basePath, ''));
}


function Combine-File {
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
                $internalFunctionPath = Join-Path -Path $file.DirectoryName -ChildPath $Matches[0];

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
            $content | Add-Content -Path $DestinationPath;
        }
    }
}


function Bundle-File {
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
    Get-Content -Path "$currentDir\PScribo.Resources.psd1" | Add-Content -Path $DestinationPath;
    Add-Content -Path $DestinationPath -Value "}`r`n";

    Combine-File -Path $Path -DestinationPath $DestinationPath -Exclude $Exclude;
    Write-Host ('  Creating bundle footer.') -ForegroundColor Cyan;
    Add-Content -Path $DestinationPath -Value "#endregion PScribo Bundle v$version";
}

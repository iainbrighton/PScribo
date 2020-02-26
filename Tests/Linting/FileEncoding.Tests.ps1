$repoRoot = (Resolve-Path "$PSScriptRoot\..\..").Path;

Describe 'Linting\FileEncoding' {

    $excludedPaths = @(
                        '.git*',
                        '.vscode',
                        'DSCResources', # We'll take the public DSC resources as-is
                        'Release',
                        '*.png',
                        '*.jpg',
                        '*.docx',
                        '*.enc',
                        '*.dll',
                        'appveyor-tools',
                        'TestResults.xml',
                        'Tests',
                        'Docs',
                        'PScriboExample.*',
                        'Lib'
                    );

    function Get-FileEncoding {
    <#
        .SYNOPSIS
            Gets file encoding.
        .DESCRIPTION
            The Get-FileEncoding function determines encoding by looking at Byte Order Mark (BOM).
            Based on port of C# code from http://www.west-wind.com/Weblog/posts/197245.aspx
        .OUTPUTS
            System.Text.Encoding
        .PARAMETER Path
            The Path of the file that we want to check.
        .PARAMETER DefaultEncoding
            The Encoding to return if one cannot be inferred.
            You may prefer to use the System's default encoding:  [System.Text.Encoding]::Default
            List of available Encodings is available here: http://goo.gl/GDtzj7
        .EXAMPLE
            # This command gets ps1 files in current directory where encoding is not ASCII
            Get-ChildItem  *.ps1 | select FullName, @{n='Encoding';e={Get-FileEncoding $_.FullName}} | where {[string]$_.Encoding -ne 'System.Text.ASCIIEncoding'}
        .EXAMPLE
            # Same as previous example but fixes encoding using set-content
            Get-ChildItem  *.ps1 | select FullName, @{n='Encoding';e={Get-FileEncoding $_.FullName}} | where {[string]$_.Encoding -ne 'System.Text.ASCIIEncoding'} | foreach {(get-content $_.FullName) | set-content $_.FullName -Encoding ASCII}
        .NOTES
            Version History
            v1.0   - 2010/08/10, Chad Miller - Initial release
            v1.1   - 2010/08/16, Jason Archer - Improved pipeline support and added detection of little endian BOMs. (http://poshcode.org/2075)
            v1.2   - 2015/02/03, VertigoRay - Adjusted to use .NET's [System.Text.Encoding Class](http://goo.gl/XQNeuc). (http://poshcode.org/5724)
        .LINK
            https://vertigion.com/2015/02/04/powershell-get-fileencoding/
    #>
        [CmdletBinding()]
        param (
            [Alias('PSPath')]
            [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
            [System.String] $Path,

            [Parameter(Mandatory = $False)]
            [System.Text.Encoding] $DefaultEncoding = [System.Text.Encoding]::ASCII
        )
        process
        {
            if ($PSVersionTable['PSEdition'] -eq 'Core')
            {
                ## PowerShell Core does not have -Encoding Byte
                [System.Byte[]] $bom = Get-Content -ReadCount 4 -TotalCount 4 -Path $Path -AsByteStream
            }
            else
            {
                [System.Byte[]] $bom = Get-Content -Encoding Byte -ReadCount 4 -TotalCount 4 -Path $Path
            }
            $hasEncoding = $false

            foreach ($encoding in [System.Text.Encoding]::GetEncodings().GetEncoding())
            {
                $preamble = $encoding.GetPreamble()
                if ($preamble)
                {
                    foreach ($i in 0..($preamble.Length -1))
                    {
                        if ($preamble[$i] -ne $bom[$i])
                        {
                            break
                        }
                        elseif ($i -eq $preamble.Length -1)
                        {
                            $hasEncoding = $encoding
                        }
                    }
                }
            }

            if (-not $hasEncoding)
            {
                Write-Warning -Message ($localized.NoFileEncodingFoundWarning -f $Path);
                $hasEncoding = $DefaultEncoding
            }

            return $hasEncoding
        } #end process
    } #end function

    function TestEncodingPath {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [System.String] $Path,

            [System.String[]] $Exclude
        )
        process
        {
            $WarningPreference = 'SilentlyContinue'
            Get-ChildItem -Path $Path -Exclude $Exclude |
                ForEach-Object {
                    if ($_ -is [System.IO.FileInfo])
                    {
                        if ($_.Name -ne 'Resolve-ProgramFilesFolder.ps1')
                        {
                            It "File '$($_.FullName.Replace($repoRoot,''))' uses UTF-8 (no BOM) encoding" {
                                $encoding = (Get-FileEncoding -Path $_.FullName -WarningAction SilentlyContinue).HeaderName
                                $encoding | Should Be 'us-ascii'
                            }
                        }
                    }
                    elseif ($_ -is [System.IO.DirectoryInfo])
                    {
                        TestEncodingPath -Path $_.FullName -Exclude $Exclude
                    }
                }
        } #end process
    } #end function

    Get-ChildItem -Path $repoRoot -Exclude $excludedPaths |
        ForEach-Object {
            TestEncodingPath -Path $_.FullName -Exclude $excludedPaths
        }
}

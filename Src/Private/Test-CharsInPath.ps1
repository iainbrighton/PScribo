Function Test-CharsInPath {

<#

    .SYNOPSIS
    PowerShell function intended to verify if in the string what is the path to file or folder are incorrect chars.

    .DESCRIPTION
    PowerShell function intended to verify if in the string what is the path to file or folder are incorrect chars.

    Exit codes

    - 0 - everything OK
    - 1 - nothing to check
    - 2 - an incorrect char found in the path part
    - 3 - an incorrect char found in the file name part
    - 4 - incorrect chars found in the path part and in the file name part

    .PARAMETER Path
    Specifies the path to an item for what path (location on the disk) need to be checked.

    The Path can be an existing file or a folder on a disk provided as a PowerShell object or a string e.g. prepared to be used in file/folder creation.

    .PARAMETER SkipCheckCharsInFolderPart
    Skip checking in the folder part of path.

    .PARAMETER SkipCheckCharsInFileNamePart
    Skip checking in the file name part of path.

    .PARAMETER SkipDividingForParts
    Skip dividing provided path to a directory and a file name.

    Used usually in conjuction with SkipCheckCharsInFolderPart or SkipCheckCharsInFileNamePart.

    .EXAMPLE

    [PS] > Test-CharsInPath -Path $(Get-Item C:\Windows\Temp\new.csv') -Verbose

    VERBOSE: The path provided as a string was devided to, directory part: C:\Windows\Temp ; file name part: new.csv
    0

    Testing existing file. Returned code means that all chars are acceptable in the name of folder and file.

    .EXAMPLE

    [PS] > Test-CharsInPath -Path "C:\newfolder:2\nowy|.csv" -Verbose

    VERBOSE: The path provided as a string was devided to, directory part: C:\newfolder:2\ ; file name part: nowy|.csv
    VERBOSE: The incorrect char | with the UTF code [124] found in FileName part
    3

    Testing the string if can be used as a file name. The returned value means that can't do to an unsupported char in the file name.

    .OUTPUTS
    Exit code as an integer number. See description section to find the exit codes descriptions.

    .LINK
    https://github.com/it-praktyk/New-OutputObject

    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech

    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, FileSystem

    REMARKS:
    # For Windows - based on the Power Tips
    # Finding Invalid File and Path Characters
    # http://community.idera.com/powershell/powertips/b/tips/posts/finding-invalid-file-and-path-characters
    # For PowerShell Core
    # https://docs.microsoft.com/en-us/dotnet/api/system.io.path.getinvalidpathchars?view=netcore-2.0
    # https://www.dwheeler.com/essays/fixing-unix-linux-filenames.html
    # [char]0 = NULL

    CURRENT VERSION
    - 0.6.1 - 2017-07-23

    HISTORY OF VERSIONS
    https://github.com/it-praktyk/New-OutputObject/CHANGELOG.md


#>

    [cmdletbinding()]
    [OutputType([System.Int32])]
    param (

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Path,
        [parameter(Mandatory = $false)]
        [switch]$SkipCheckCharsInFolderPart,
        [parameter(Mandatory = $false)]
        [switch]$SkipCheckCharsInFileNamePart,
        [parameter(Mandatory = $false)]
        [switch]$SkipDividingForParts

    )

    BEGIN {

        If (($PSVersionTable.ContainsKey('PSEdition')) -and ($PSVersionTable.PSEdition -eq 'Core') -and $IsLinux) {

            #[char]0 = NULL
            $PathInvalidChars = [char]0

            $FileNameInvalidChars = @([char]0, '/')

            $PathSeparators = @('/')

        }
        Elseif (($PSVersionTable.ContainsKey('PSEdition')) -and ($PSVersionTable.PSEdition -eq 'Core') -and $IsMacOS) {

            $PathInvalidChars = [char]58

            $FileNameInvalidChars = [char]58

            $PathSeparators = @('/')

        }
        #Windows
        Else {

            $PathInvalidChars = [System.IO.Path]::GetInvalidPathChars() #36 chars

            $FileNameInvalidChars = [System.IO.Path]::GetInvalidFileNameChars() #41 chars

            #$FileOnlyInvalidChars = @(':', '*', '?', '\', '/') #5 chars - as a difference

            $PathSeparators = @('/','\')

        }

        $IncorectCharFundInPath = $false

        $IncorectCharFundInFileName = $false

        $NothingToCheck = $true

    }

    END {

        [String]$DirectoryPath = ""

        [String]$FileName = ""

        $PathType = ($Path.GetType()).Name

        If (@('DirectoryInfo', 'FileInfo') -contains $PathType) {

            If (($SkipCheckCharsInFolderPart.IsPresent -and $PathType -eq 'DirectoryInfo') -or ($SkipCheckCharsInFileNamePart.IsPresent -and $PathType -eq 'FileInfo')) {

                Return 1

            }
            ElseIf ($PathType -eq 'DirectoryInfo') {

                [String]$DirectoryPath = $Path.FullName

            }

            elseif ($PathType -eq 'FileInfo') {

                [String]$DirectoryPath = $Path.DirectoryName

                [String]$FileName = $Path.Name

            }

        }

        ElseIf ($PathType -eq 'String') {

            If ( $SkipDividingForParts.IsPresent -and $SkipCheckCharsInFolderPart.IsPresent ) {

                $FileName = $Path

            }
            ElseIf ( $SkipDividingForParts.IsPresent -and $SkipCheckCharsInFileNamePart.IsPresent  ) {

                $DirectoryPath = $Path

            }
            Else {

                #Convert String to Array of chars
                $PathArray = $Path.ToCharArray()

                $PathLength = $PathArray.Length

                For ($i = ($PathLength-1); $i -ge 0; $i--) {

                    If ($PathSeparators -contains $PathArray[$i]) {

                        [String]$DirectoryPath = [String]$Path.Substring(0, $i +1)

                        break

                    }

                }

                If ([String]::IsNullOrEmpty($DirectoryPath)) {

                    [String]$FileName = [String]$Path

                }
                Else {

                    [String]$FileName = $Path.Replace($DirectoryPath, "")

                }

            }

        }
        Else {

            [String]$MessageText = "Input object {0} can't be tested" -f ($Path.GetType()).Name

            Throw $MessageText

        }

        [String]$MessageText = "The path provided as a string was divided to: directory part: {0} ; file name part: {1} ." -f $DirectoryPath, $FileName

        Write-Verbose -Message $MessageText

        If ($SkipCheckCharsInFolderPart.IsPresent -and $SkipCheckCharsInFileNamePart.IsPresent) {

            Return 1

        }

        If (-not ($SkipCheckCharsInFolderPart.IsPresent) -and -not [String]::IsNullOrEmpty($DirectoryPath)) {

            $NothingToCheck = $false

            foreach ($Char in $PathInvalidChars) {

                If ($DirectoryPath.ToCharArray() -contains $Char) {

                    $IncorectCharFundInPath = $true

                    [String]$MessageText = "The incorrect char {0} with the UTF code [{1}] found in the Path part." -f $Char, $([int][char]$Char)

                    Write-Verbose -Message $MessageText

                }

            }

        }

        If (-not ($SkipCheckCharsInFileNamePart.IsPresent) -and -not [String]::IsNullOrEmpty($FileName)) {

            $NothingToCheck = $false

            foreach ($Char in $FileNameInvalidChars) {

                If ($FileName.ToCharArray() -contains $Char) {

                    $IncorectCharFundInFileName = $true

                    [String]$MessageText = "The incorrect char {0} with the UTF code [{1}] found in FileName part." -f $Char, $([int][char]$Char)

                    Write-Verbose -Message $MessageText

                }

            }

        }

        If ($IncorectCharFundInPath -and $IncorectCharFundInFileName) {

            Return 4

        }
        elseif ($NothingToCheck) {

            Return 1

        }

        elseif ($IncorectCharFundInPath) {

            Return 2

        }

        elseif ($IncorectCharFundInFileName) {

            Return 3

        }
        Else {

            Return 0

        }

    }

}

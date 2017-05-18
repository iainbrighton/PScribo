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
    # Based on the Power Tips
    # Finding Invalid File and Path Characters
    # http://powershell.com/cs/blogs/tips/archive/2016/04/20/finding-invalid-file-and-path-characters.aspx

    CURRENT VERSION
    - 0.5.4 - 2017-05-16

    HISTORY OF VERSIONS
    https://github.com/it-praktyk/New-OutputObject/VERSIONS.md


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
        [switch]$SkipCheckCharsInFileNamePart

    )

    BEGIN {

        $PathInvalidChars = [System.IO.Path]::GetInvalidPathChars() #36 chars

        $FileNameInvalidChars = [System.IO.Path]::GetInvalidFileNameChars() #41 chars

        #$FileOnlyInvalidChars = @(':', '*', '?', '\', '/') #5 chars - as a difference

        $IncorectCharFundInPath = $false

        $IncorectCharFundInFileName = $false

        $NothingToCheck = $true

    }

    PROCESS {

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

            [String]$DirectoryPath = ""

            #Convert String to Array of chars
            $PathArray = $Path.ToCharArray()

            $PathLength = $PathArray.Length

            #It should be removed in a future and for loop changed to $i--
            #Reverse PathArray to speedup of finding the folder indicator chars '\' or '/'
            [array]::Reverse($PathArray)

            For ($i = 0; $i -lt $PathLength; $i++) {

                If (@('\', '/') -contains $PathArray[$i]) {

                    [String]$DirectoryPath = [String]$Path.Substring(0, $($PathArray.Length - $i))

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

    }

    END {

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

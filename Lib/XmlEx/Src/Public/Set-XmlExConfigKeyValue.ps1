function Set-XmlExConfigKeyValue {
<#
    .SYNOPSIS
        Ensures an Xml document contains a Xml element/attribute.
#>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Path')]
    param (
        # Specifies a path to one or more locations. Wildcards are permitted.
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Path', ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [System.String[]] $Path,

        # Specifies a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is
        # used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters,
        # enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any
        # characters as escape sequences.
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'LiteralPath', ValueFromPipelineByPropertyName)]
        [Alias('PSPath')]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $LiteralPath,

        ## Specifies the XPath location of the parent Xml element within the document to update.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $XPath,

        ## Specifies the name of the Xml element or Xml attribute to update.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Name,

        ## Specifies the value of the Xml element or Xml attribute to update.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $Value,

        ## Specifies the target is an attribute. Defaults to an element.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $IsAttribute,

        ## Specifies the target Xml document file should be created if it does not exist.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $Force
    )
    begin {

        [ref] $null = $PSBoundParameters.Remove('Path');
        [ref] $null = $PSBoundParameters.Remove('LiteralPath');
        $paths = @();

    }
    process {

        if ($PSCmdlet.ParameterSetName -eq 'Path') {

            foreach ($filePath in $Path) {

                if (-not $Force) {

                    if (-not (Test-Path -Path $filePath)) {

                        $errorMessage = $localized.CannotFindPathError -f $filePath;
                        $ex = New-Object System.Management.Automation.ItemNotFoundException $errorMessage;
                        $category = [System.Management.Automation.ErrorCategory]::ObjectNotFound;
                        $errorRecord = New-Object System.Management.Automation.ErrorRecord $ex, 'PathNotFound', $category, $filePath;
                        $PSCmdlet.WriteError($errorRecord);
                        continue;
                    }

                    # Resolve any wildcards that might be in the path
                    $provider = $null;
                    $paths += $psCmdlet.SessionState.Path.GetResolvedProviderPathFromPSPath($filePath, [ref] $provider);

                }
                else {

                    $paths += $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($filePath);

                }

            } #end foreach path
        }
        else {

            foreach ($filePath in $LiteralPath) {

                if (-not $Force) {

                    if (-not (Test-Path -LiteralPath $filePath)) {

                        $errorMessage = $localized.CannotFindPathError -f $filePath;
                        $ex = New-Object System.Management.Automation.ItemNotFoundException $errorMessage;
                        $category = [System.Management.Automation.ErrorCategory]::ObjectNotFound;
                        $errorRecord = New-Object System.Management.Automation.ErrorRecord $ex, 'PathNotFound', $category, $filePath;
                        $PSCmdlet.WriteError($errorRecord);
                        continue;
                    }

                }

                # Resolve any relative paths
                $paths += $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($filePath);

            } #end foreach literal path
        }

    } #end process
    end {

        foreach ($filePath in $paths) {

            Write-Verbose -Message ($localized.ProcessingDocument -f $filePath);
            $xmlDocument = Assert-XmlExFilePath -Path $filePath -Ensure 'Present';;
            Assert-XmlExXPath -XmlDocument $xmlDocument @PSBoundParameters;

            $verboseMessage = $localized.SavingDocument -f $filePath;
            $warningConfirmationMessage = $localized.ShouldProcessWarning;
            $warningDescriptionMessage = $localized.ShouldProcessOperationWarning -f 'Save', $filePath;

            if ($Force -or ($PSCmdlet.ShouldProcess($verboseMessage, $warningConfirmationMessage, $warningDescriptionMessage))) {

                $xmlDocument.Save($filePath);
            }
        }

    } #end
} #end function

# SIG # Begin signature block
# MIIX1gYJKoZIhvcNAQcCoIIXxzCCF8MCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMgorJ7tt90pfruefEtpDHi60
# kGGgghMJMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9v
# dCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNp
# Z25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4R
# r2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrw
# nIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnC
# wlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8
# y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM
# 0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6f
# pjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1Ud
# DwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGsw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcw
# AoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Um9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBP
# BgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoK
# o6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8w
# DQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+
# C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119E
# efM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR
# 4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4v
# cn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwH
# gfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggU4MIIEIKADAgEC
# AhAPxQCJrE9ObGzkCRS7EwyyMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25p
# bmcgQ0EwHhcNMTcwNTI2MDAwMDAwWhcNMTkwOTI3MTIwMDAwWjB1MQswCQYDVQQG
# EwJHQjETMBEGA1UECBMKR2Fyc2luZ3RvbjEPMA0GA1UEBxMGT3hmb3JkMR8wHQYD
# VQQKExZWaXJ0dWFsIEVuZ2luZSBMaW1pdGVkMR8wHQYDVQQDExZWaXJ0dWFsIEVu
# Z2luZSBMaW1pdGVkMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnB1O
# DV2jw/aMIUWnD9f9RCbAoiJ8LQcznYo42P22YEi6g7QY+kKmAzGgEhbsE4UVuGWS
# el4y6FxGWq51SK5P/gqgZgzyP8FkIUzLxsCrtx9OBnsPPeL+/An5CpcsKsl2lCSz
# NMwcz16hjcE0vCLio1NOcwvfO65qdNT2gRIEnIYhRX88dG3V30BH2aKWG5X9t1IW
# RmozjZ8I7iLEoWFJWQSuICSGyvyiPqnXF3nxdroE8O4fc1U90x5qefX0RlwKeq47
# UFuI0Y/59pB3/jss5BYvAXp3g+6EKlDwgW1a/MLVsLQbdzlALFUv5YxEqkXA8IEM
# xpwgBjm117SmyZ98QQIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6Xq
# cQPAYPkt9mV1DlgwHQYDVR0OBBYEFL5NkOqMm0S8AyuXI1EZIdoK9DD/MA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Ax
# hi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1j
# cy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUF
# BwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4G
# CCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRT
# SEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkq
# hkiG9w0BAQsFAAOCAQEAQC8qzSz1bIoEqbjDx3VtYDjtUjuFEVDFYi9+vREl6jM+
# iqOiNiwCXUkbxGTuDrWW9I1YOn8a7SCCYapZ+T0G3RMa+rQHXFYKbYTmXC3C41Cd
# MQzZn4wTuGRNQLTgNSuclwMnNmFVe7K5S/0Dv+GaLSKuRLAxpcPxeTtmRZIIBXF7
# wwRS0+V28jB9VyeSOdqsPIFYf5GSfu7KcIhmNQ/DUroulaS5JIrPUhwkf1LZMm0B
# /0adpaPbFy95M1emij96rrgy2hX8N/FrWBh13/81V6NO3b8mhCfjqb632dG4EUTi
# FXDvqP2hpWw0nO/pFZsMsEK88eiV93XDDEG/MjAApzGCBDcwggQzAgEBMIGGMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBDb2RlIFNpZ25pbmcgQ0ECEA/FAImsT05sbOQJFLsTDLIwCQYFKw4DAhoFAKB4
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFKbfrmBGJ+X0fwpRUJPdlzTVqbG6MA0GCSqGSIb3DQEBAQUABIIBACyP7VRV
# 5D+4iAjMqAX88fNeQb8qqxFAos9hLGT0DCnjwHroK57+BqU3hj60pKtw1PeuMV5c
# mZUPpEzl/0Mp0mnPxgkrAk9nmzdZDUiAA7ixqbnnHU31NC9U5WWa3aso1oXPzRyW
# Ff0kO0n42asUkuEPgGpWF6gHxlD2f5EH0sDQpz3vxXVHCAE5kWwKWCKsaAP0ljIu
# CHuHCQnR3NFpGthUqLblDe1onXivzFmotUInVlf1m7+lL/yLoqqEJV8x/vaPLvZh
# hbh7Wfap+hqXN3g0JUJdHKAkQDqgMpAqT8j62s+8/RZkrl+uPdQW7DYNbuJL9qCt
# fp2XLcXcIjj0DU+hggILMIICBwYJKoZIhvcNAQkGMYIB+DCCAfQCAQEwcjBeMQsw
# CQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNV
# BAMTJ1N5bWFudGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMgIQDs/0
# OMj+vzVuBNhqmBsaUDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTkwNjE4MTY1NzE4WjAjBgkqhkiG9w0BCQQx
# FgQUI6kwrMYDAghXCYcDJQJM8mR861AwDQYJKoZIhvcNAQEBBQAEggEAhPrawtD1
# El+1bdq5DUXCIq+Mj3rrrawmyGSd4O1I1OMtAiy9utPpObE6NCo0Rtq9g4lvjsjc
# VvM+9GpVqXfhSVY0+Rr8BBj9nyVmADk4pS5cayShsyHM4kemOzmfBcCwL9M+ERLq
# s8bSQ+NNqk1EpfMy+dy8Ml3rSTptcFhvLSmzhvC6BzFmZSHjM4fNgSPQSYV9Z+D8
# DZwtwKwYbOrgkdVRoXnZR8/pA2gaU7o0Ren6GpbJ/5mkN3rRc/NiSZInlpjW46MJ
# Al/WRRIOhTzgb6x81eSPZX9615Ik81a8kAXbf+tVc16/2auojJ1kOsK4QBltr0Xv
# YLIkaWPkiaqzjw==
# SIG # End signature block

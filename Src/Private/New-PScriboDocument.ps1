function New-PScriboDocument
{
<#
    .SYNOPSIS
        Initializes a new PScript document object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseLiteralInitializerForHashtable','')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        ## PScribo document name
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## PScribo document Id
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [System.String] $Id = $Name.Replace(' ','')
    )
    begin
    {
        if ($(Test-CharsInPath -Path $Name -SkipCheckCharsInFolderPart -Verbose:$false) -eq 3 )
        {
            throw -Message ($localized.IncorrectCharsInName)
        }
    }
    process
    {
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Name)
        $typeName = 'PScribo.Document'
        $pscriboDocument = [PSCustomObject] @{
            Id                = $Id.ToUpper()
            Type              = $typeName
            Name              = $Name
            Sections          = New-Object -TypeName System.Collections.ArrayList
            Options           = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase)
            Properties        = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase)
            Styles            = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase)
            TableStyles       = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase)
            DefaultStyle      = $null
            DefaultTableStyle = $null
            Header            = [PSCustomObject] @{
                                    HasFirstPageHeader = $false
                                    HasDefaultHeader   = $false
                                    FirstPageHeader    = $null
                                    DefaultHeader      = $null
                                }
            Footer            = [PSCustomObject] @{
                                    HasFirstPageFooter = $false
                                    HasDefaultFooter   = $false
                                    FirstPageFooter    = $null
                                    DefaultFooter      = $null
                                }
            TOC               = New-Object -TypeName System.Collections.ArrayList
        }
        $defaultDocumentOptionParams = @{
            MarginTopAndBottom = 72
            MarginLeftAndRight = 54
            Orientation        = 'Portrait'
            PageSize           = 'A4'
            DefaultFont        = 'Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif'
        }
        DocumentOption @defaultDocumentOptionParams -Verbose:$false

        ## Set "default" styles
        Style -Name Normal -Size 11 -Default -Verbose:$false
        Style -Name Title -Size 28 -Color '0072af' -Verbose:$false
        Style -Name TOC -Size 16 -Color '0072af' -Hide -Verbose:$false
        Style -Name 'Heading 1' -Size 16 -Color '0072af' -Verbose:$false
        Style -Name 'Heading 2' -Size 14 -Color '0072af' -Verbose:$false
        Style -Name 'Heading 3' -Size 12 -Color '0072af' -Verbose:$false
        Style -Name 'Heading 4' -Size 11 -Color '2f5496' -Italic -Verbose:$false
        Style -Name 'Heading 5' -Size 11 -Color '2f5496' -Verbose:$false
        Style -Name 'Heading 6' -Size 11 -Color '1f3763' -Verbose:$false
        Style -Name TableDefaultHeading -Size 11 -Color 'fff' -BackgroundColor '4472c4' -Bold -Verbose:$false
        Style -Name TableDefaultRow -Size 11 -Verbose:$false
        Style -Name TableDefaultAltRow -Size 11 -BackgroundColor 'd0ddee' -Verbose:$false
        Style -Name Caption -Size 11 -Italic -Verbose:$false
        $tableDefaultStyleParams = @{
            Id                = 'TableDefault'
            BorderWidth       = 1
            BorderColor       = '2a70be'
            HeaderStyle       = 'TableDefaultHeading'
            RowStyle          = 'TableDefaultRow'
            AlternateRowStyle = 'TableDefaultAltRow'
            CaptionStyle      = 'Caption'
        }
        TableStyle @tableDefaultStyleParams -Default -Verbose:$false
        return $pscriboDocument
    }
}

[CmdletBinding()]
param (
    [System.String[]] $Format = @('Html','Word'),
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

<# The document name is used in the file output #>
$document = Document 'PScribo Demo 2' -Verbose {
    <#  Enforce uppercase section headers/names
        Enable automatic section numbering
        Set the page size to US Letter with 0.5inch margins #>
    DocumentOption -ForceUppercaseSection -EnableSectionNumbering -PageSize Letter -Margin 36
    BlankLine -Count 20
    Paragraph 'PScribo Demo 2' -Style Title
    BlankLine -Count 20
    PageBreak
    TOC -Name 'Table of Contents'
    PageBreak

    <# WARNING:
        Microsoft Word will include paragraphs styled with 'Heading*' style names to the TOC.
        To avoid this, define an identical style with a name not beginning with 'Heading'!
    #>
    Paragraph -Style Heading1 'This is Heading 1'
    Paragraph -Style Heading2 'This is Heading 2'
    Paragraph -Style Heading3 'This is Heading 3'
    Paragraph 'This is a regular line of text indented 0 tab stops'
    Paragraph -Tabs 1 'This is a regular line of text indented 1 tab stops. This text should not be displayed as a hanging indent, e.g. not just the first line of the paragraph indented.'
    Paragraph -Tabs 2 'This is a regular line of text indented 2 tab stops'
    Paragraph -Tabs 3 'This is a regular line of text indented 3 tab stops'
    Paragraph 'This is a regular line of text in the default font in italics' -Italic
    Paragraph 'This is a regular line of text in the default font in bold' -Bold
    Paragraph 'This is a regular line of text in the default font in bold italics' -Bold -Italic
    Paragraph 'This is a regular line of text in the default font in 14 point' -Size 14
    Paragraph 'This is a regular line of text in Courier New font' -Font 'Courier New'
    Paragraph "This is a regular line of text indented 0 tab stops with the computer name as data: $env:COMPUTERNAME"
    Paragraph "This is a regular line of text indented 0 tab stops with the computer name as data in bold: $env:COMPUTERNAME" -Bold
    Paragraph "This is a regular line of text indented 0 tab stops with the computer name as data in bold italics: $env:COMPUTERNAME" -Bold -Italic
    Paragraph "This is a regular line of text indented 0 tab stops with the computer name as data in 14 point bold italics: $env:COMPUTERNAME" -Bold -Italic -Size 14
    Paragraph "This is a regular line of text indented 0 tab stops with the computer name as data in 8 point Courier New bold italics: $env:COMPUTERNAME" -Bold -Italic -Size 8 -Font 'Courier New'

    $services = Get-CimInstance -ClassName Win32_Service | Select-Object -Property DisplayName, State, StartMode | Sort-Object -Property DisplayName
    <# Add a custom style for highlighting table cells/rows #>
    Style -Name 'Stopped Service' -Color White -BackgroundColor Firebrick -Bold

    <#  Sections provide an easy way of creating a document structure and can support section
        numbering (if enabled with the DocumentOption -EnableSectionNumbering parameter. You don't
        need to worry about the numbers as PScribo will figure this out. #>
    Section -Style Heading1 'Standard-Style Tables' {
        Section -Style Heading2 'Autofit Width Autofit Cell No Highlighting' {
            Paragraph -Style Heading3 'Example of an autofit table width, autofit contents and no cell highlighting.'
            Paragraph "Services ($($services.Count) Services found):"
            $services | Table -Columns DisplayName,State,StartMode -Headers 'Display Name','Status','Startup Type' -Width 0
        }
        PageBreak
        Section -Style Heading2 'Full Width Autofit Cell Highlighting' {
            Paragraph -Style Heading3 'Example of a full width table with autofit columns and individual cell highlighting.'
            Paragraph "Services ($($services.Count) Services found):"
            <# Highlight individual cells with "StoppedService" style where state = stopped and startup = auto #>
            $stoppedAutoServicesCell = $services.Clone()
            $stoppedAutoServicesCell | Where-Object { $_.State -eq 'Stopped' -and $_.StartMode -eq 'Auto'} | Set-Style -Property State -Style StoppedService
            $stoppedAutoServicesCell | Table -Columns DisplayName,State,StartMode -Headers 'Display Name','Status','Startup Type' -Tabs 1
        }

        PageBreak
        Section -Style Heading2 'Full Width Fixed Row Highlighting' {
            Paragraph -Style Heading3 'Example of a full width table with fixed columns widths and full row highlighting.'
            Paragraph "Services ($($services.Count) Services found)"
            <# Highlight an entire row with the "StoppedService" style where state = stopped and startup = auto #>
            $stoppedAutoServicesRow = $services.Clone()
            $stoppedAutoServicesRow | Where-Object { $_.State -eq 'Stopped' -and $_.StartMode -eq 'Auto' } | Set-Style -Style StoppedService
            $stoppedAutoServicesRow | Table -Columns DisplayName,State,StartMode -ColumnWidths 70,15,15 -Headers 'Display Name','Status','Startup Type'
        }
    }

    PageBreak
    $listServices = Get-Service | Select-Object -Property Name,CanPauseAndContinue,CanShutdown,CanStop,DisplayName,ServiceName,ServiceType,Status -First 2
    $listServices | Set-Style -Style StoppedService -Property ServiceType -Verbose

    Section -Style Heading1 'List-Style Tables' {
        Section -Style Heading2 'Autofit Width Autofit Cell No Highlighting' {
            $listServices | Table -Name 'AutofitWidth-AutofitCell-NoHighlighting' -List -Width 0
        }

        Section -Style Heading2 'Fixed Width Autofit Cell No Highlighting' {
            foreach ($listService in $listServices) {
                Paragraph "Service $($listService.Name):"
                $listService | Table -Name "FixedWidth-AutofitCell-$($listService.Name)" -List -Width 60
            }
        }

        Section -Style Heading2 'Fixed Width Fixed Cell No Highlighting 1' {
            foreach ($listService in $listServices) {
                Paragraph "Service $($listService.Name):" -Bold
                $listService | Table -Name "FixedWidth-FixedCell-$($listService.Name)1" -List -ColumnWidths 25,75 -Width 60
            }
        }

        Section -Style Heading1 'Fixed Width Fixed Cell No Highlighting 2' {
            $listServices |  Table -Name "FixedWidth-FixedCell-$($listService.Name)2" -List -ColumnWidths 45,55 -Width 60 -Tabs 2
        }
    }
}
<#  Generate 'PScribo Demo 2.docx' and 'PScribo Demo 2.html' files. Other supported formats include 'Text' and 'Xml' #>
$document | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

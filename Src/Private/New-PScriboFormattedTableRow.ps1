function New-PScriboFormattedTableRow
{
<#
    .SYNOPSIS
        Creates a formatted table row for plugin output/rendering.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Management.Automation.PSObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.String] $TableStyle,

        [Parameter(ValueFromPipeline)]
        [AllowNull()]
        [System.String] $Style = $null,

        [Parameter(ValueFromPipeline, ParameterSetName = 'Header')]
        [System.Management.Automation.SwitchParameter] $IsHeaderRow,

        [Parameter(ValueFromPipeline, ParameterSetName = 'Row')]
        [System.Management.Automation.SwitchParameter] $IsAlternateRow
    )
    process
    {
        if (-not ([System.String]::IsNullOrEmpty($Style)))
        {
            ## Use the explictit style
            $IsStyleInherited = $false
        }
        elseif ($IsHeaderRow)
        {
            $Style = $document.TableStyles[$TableStyle].HeaderStyle
            $IsStyleInherited = $true
        }
        elseif ($IsAlternateRow)
        {
            $Style = $document.TableStyles[$TableStyle].AlternateRowStyle
            $IsStyleInherited = $true
        }
        else
        {
            $Style = $document.TableStyles[$TableStyle].RowStyle
            $IsStyleInherited = $true
        }

        return [PSCustomObject] @{
            Style            = $Style;
            IsStyleInherited = $IsStyleInherited;
            Cells            = New-Object -TypeName System.Collections.ArrayList;
        }
    }
}

function OutTextTable
{
<#
    .SYNOPSIS
        Output formatted text table.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Table
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboTextOption
        }
    }
    process
    {
        ## Use the current output buffer width
        if ($options.TextWidth -eq 0)
        {
            $options.TextWidth = $Host.UI.RawUI.BufferSize.Width -1
        }
        $tableWidth = $options.TextWidth - ($Table.Tabs * 4)
        if ($Table.IsList)
        {
            $tableText = ($Table.Rows |
                Select-Object -Property * -ExcludeProperty '*__Style' |
                    Format-List | Out-String -Width $tableWidth).Trim()
        }
        else
        {
            ## Don't trim tabs for table headers
            ## Tables set to AutoSize as otherwise rendering is different between PoSh v4 and v5
            $tableText = ($Table.Rows |
                            Select-Object -Property * -ExcludeProperty '*__Style' |
                                Format-Table -Wrap -AutoSize |
                                    Out-String -Width $tableWidth).Trim("`r`n")
        }
        $tableText = ConvertTo-IndentedString -InputObject $tableText -Tabs $Table.Tabs
        return ('{0}{1}' -f [System.Environment]::NewLine, $tableText)
    }
}

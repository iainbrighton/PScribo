function Out-HtmlStyle
{
<#
    .SYNOPSIS
        Generates an in-line HTML CSS stylesheet from a PScribo document styles and table styles.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        ## PScribo document styles
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Collections.Hashtable] $Styles,

        ## PScribo document tables styles
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.Collections.Hashtable] $TableStyles,

        ## Suppress page layout styling
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $NoPageLayoutStyle
    )
    process
    {
        $stylesBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [ref] $null = $stylesBuilder.AppendLine('<style type="text/css">')
        $pageWidth = $Document.Options['PageWidth']
        $pageHeight = $Document.Options['PageHeight']
        if (-not $NoPageLayoutStyle)
        {
            ## Add HTML page layout styling options, e.g. when emailing HTML documents
            [ref] $null = $stylesBuilder.AppendLine('html { height: 100%; -webkit-background-size: cover; -moz-background-size: cover; -o-background-size: cover; background-size: cover; background: #e6e6e6; }')
            [ref] $null = $stylesBuilder.Append("page { background: white; display: block; margin-top: 1rem; margin-left: auto; margin-right: auto; margin-bottom: 1rem; ")
            [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }')
            [ref] $null = $stylesBuilder.AppendLine('@media print { body, page { margin: 0; box-shadow: 0; } }')
            [ref] $null = $stylesBuilder.AppendLine('hr { margin-top: 1.0rem; }')
            [ref] $null = $stylesBuilder.Append(" .portrait { background: white; width: $($pageWidth)mm; display: block; margin-top: 1rem; margin-left: auto; margin-right: auto; margin-bottom: 1rem; position: relative; ")
            [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }')
            [ref] $null = $stylesBuilder.Append(" .landscape { background: white; width: $($pageHeight)mm; display: block; margin-top: 1rem; margin-left: auto; margin-right: auto; margin-bottom: 1rem; position: relative; ")
            [ref] $null = $stylesBuilder.AppendLine('border-style: solid; border-width: 1px; border-color: #c6c6c6; }')
        }

        foreach ($style in $Styles.Keys)
        {
            ## Build style
            $htmlStyle = Get-HtmlStyle -Style $Styles[$style]
            [ref] $null = $stylesBuilder.AppendFormat(' .{0} {{{1} }}', $Styles[$style].Id, $htmlStyle).AppendLine()
        }

        foreach ($tableStyle in $TableStyles.Keys)
        {
            $tStyle = $TableStyles[$tableStyle]
            $tableStyleId = $tStyle.Id.ToLower()
            $htmlTableStyle = Get-HtmlTableStyle -TableStyle $tStyle
            $htmlHeaderStyle = Get-HtmlStyle -Style $Styles[$tStyle.HeaderStyle]
            $htmlRowStyle = Get-HtmlStyle -Style $Styles[$tStyle.RowStyle]
            $htmlAlternateRowStyle = Get-HtmlStyle -Style $Styles[$tStyle.AlternateRowStyle]
            ## Generate table style
            [ref] $null = $stylesBuilder.AppendFormat(' table.{0} {{{1} }}', $tableStyleId, $htmlTableStyle).AppendLine()
            [ref] $null = $stylesBuilder.AppendFormat(' table.{0} th {{{1}{2} }}', $tableStyleId, $htmlHeaderStyle, $htmlTableStyle).AppendLine()
            [ref] $null = $stylesBuilder.AppendFormat(' table.{0} td {{{1} }}', $tableStyleId,  $htmlTableStyle).AppendLine()
            [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(odd) {{{1}{2} }}', $tableStyleId, $htmlRowStyle, $htmlTableStyle).AppendLine()
            [ref] $null = $stylesBuilder.AppendFormat(' table.{0} tr:nth-child(even) {{{1}{2} }}', $tableStyleId, $htmlAlternateRowStyle, $htmlTableStyle).AppendLine()
        }

        [ref] $null = $stylesBuilder.AppendLine('</style>')
        return $stylesBuilder.ToString().TrimEnd()
    }
}

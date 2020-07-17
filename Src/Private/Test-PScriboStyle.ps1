function Test-PScriboStyle
{
<#
    .SYNOPSIS
        Tests whether a style has been defined.
#>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name
    )
    process
    {
        return $PScriboDocument.Styles.ContainsKey($Name)
    }
}

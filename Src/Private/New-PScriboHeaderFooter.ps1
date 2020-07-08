function New-PScriboHeaderFooter
{
<#
    .SYNOPSIS
        Initializes a new PScribo header/footer object.

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','Footer')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param
    (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Header')]
        [System.Management.Automation.SwitchParameter] $Header,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Footer')]
        [System.Management.Automation.SwitchParameter] $Footer
    )
    begin
    {
        ## Ignores dot-sourced script blocks, i.e. style scripts
        $psCallStack = Get-PSCallStack | Where-Object { $_.FunctionName -ne '<ScriptBlock>' }
        if ($psCallStack[2].FunctionName -ne 'Document<Process>')
        {
            throw $localized.HeaderFooterDocumentRootError
        }
    }
    process
    {
        $pscriboHeaderFooter = [PSCustomObject] @{
            Type     = if ($Header) { 'PScribo.Header' } else { 'PScribo.Footer' }
            Sections = (New-Object -TypeName System.Collections.ArrayList)
        }
        return $pscriboHeaderFooter
    }
}

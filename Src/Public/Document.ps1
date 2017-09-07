function Document {
<#
    .SYNOPSIS
        Initializes a new PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## PScribo document name
        [Parameter(Mandatory, Position = 0)]
        [System.String] $Name,

        ## PScribo document DSL script block containing Section, Paragraph and/or Table etc. commands.
        [Parameter(Position = 1)]
        [System.Management.Automation.ScriptBlock] $ScriptBlock = $(throw $localized.NoScriptBlockProvidedError),

        ## PScribo document Id
        [Parameter()]
        [System.String] $Id = $Name.Replace(' ','')
    )
    begin {

        $pluginName = 'Document';
        <#! Document.Internal.ps1 !#>

    } #end begin
    process {

        $stopwatch = [Diagnostics.Stopwatch]::StartNew();
        $pscriboDocument = New-PScriboDocument -Name $Name -Id $Id;

        ## Call the Document script block
        foreach ($result in & $ScriptBlock) {

            [ref] $null = $pscriboDocument.Sections.Add($result);
        }

        Invoke-PScriboSection;

        WriteLog -Message ($localized.DocumentProcessingCompleted -f $pscriboDocument.Name);
        $stopwatch.Stop();
        WriteLog -Message ($localized.TotalProcessingTime -f $stopwatch.Elapsed.TotalSeconds);
        return $pscriboDocument;

    } #end process
} #end function Document

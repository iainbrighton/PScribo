function Export-Document {
<#
    .SYNOPSIS
        Exports a PScribo document object to one or more output formats.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock','')]
    [OutputType([System.IO.FileInfo])]
    param
    (
        ## PScribo document object
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ## Output file path
        [Parameter(Position = 2, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path = (Get-Location -PSProvider FileSystem),

        ## PScribo document export option
        [Parameter(Position = 3, ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options,

        [Parameter(Position = 4, ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    )
    DynamicParam
    {
        ## Adds a dynamic -Format parameter that returns the available plugins
        $parameterAttribute = New-Object -TypeName 'System.Management.Automation.ParameterAttribute'
        $parameterAttribute.ParameterSetName = '__AllParameterSets'
        $parameterAttribute.Mandatory = $true
        $parameterAttribute.Position = 1

        $attributeCollection = New-Object -TypeName 'System.Collections.ObjectModel.Collection[System.Attribute]'
        $attributeCollection.Add($parameterAttribute)
        $validateSetAttribute = New-Object -TypeName 'System.Management.Automation.ValidateSetAttribute' -ArgumentList (Get-PScriboPlugin)
        $attributeCollection.Add($validateSetAttribute)

        $runtimeParameter = New-Object -TypeName 'System.Management.Automation.RuntimeDefinedParameter' -ArgumentList @('Format', [System.String[]], $attributeCollection)
        $runtimeParameterDictionary = New-Object -TypeName 'System.Management.Automation.RuntimeDefinedParameterDictionary'
        $runtimeParameterDictionary.Add('Format', $runtimeParameter)
        return $runtimeParameterDictionary
    }
    begin
    {
        try { $Path = Resolve-Path $Path -ErrorAction SilentlyContinue; }
        catch { }

        if ( $(Test-CharsInPath -Path $Path -SkipCheckCharsInFileNamePart -Verbose:$false) -eq 2 )
        {
            throw $localized.IncorrectCharsInPathError
        }

        if (-not (Test-Path $Path -PathType Container))
        {
            ## Check $Path is a directory
            throw ($localized.InvalidDirectoryPathError -f $Path)
        }
    }
    process
    {
        foreach ($f in $PSBoundParameters['Format'])
        {
            Write-PScriboMessage -Message ($localized.DocumentInvokePlugin -f $f) -Plugin 'Export';

            ## Dynamically generate the output format function name
            $outputFormat = 'Out-{0}Document' -f $f
            $outputParams = @{
                Document = $Document
                Path = $Path
            }
            if ($PSBoundParameters.ContainsKey('Options'))
            {
                $outputParams['Options'] = $Options
            }

            $fileInfo = & $outputFormat @outputParams
            Write-PScriboMessage -Message ($localized.DocumentExportPluginComplete -f $f) -Plugin 'Export'

            if ($PassThru)
            {
                Write-Output -InputObject $fileInfo
            }
        }
    }
}

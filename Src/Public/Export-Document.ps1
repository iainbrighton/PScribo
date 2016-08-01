function Export-Document {
<#
    .SYNOPSIS
        Exports a PScribo document object to one or more output formats.
#>
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo])]
    param (
        ## PScribo document object
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object] $Document,

        ## Output formats
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.String[]] $Format,

        ## Output file path
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path = (Get-Location -PSProvider FileSystem),

        ## PScribo document export option
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        try { $Path = Resolve-Path $Path -ErrorAction SilentlyContinue; }
        catch { }

        if (-not (Test-Path $Path -PathType Container)) {
            ## Check $Path is a directory
            throw ($localized.InvalidDirectoryPathError -f $Path);
        }

    }
    process {

        foreach ($f in $Format) {
            WriteLog -Message ($localized.DocumentInvokePlugin -f $f) -Plugin 'Export';
            ## Call specified output plugin
            #try {
                ## Dynamically generate the output format function name
                $outputFormat = 'Out{0}' -f $f;
                if ($PSBoundParameters.ContainsKey('Options')) {
                    & $outputFormat -Document $Document -Path $Path -Options $Options; # -ErrorAction Stop;
                }
                else {
                    & $outputFormat -Document $Document -Path $Path; # -ErrorAction Stop;
                }
            #}
            #catch [System.Management.Automation.CommandNotFoundException] {
            #    Write-Warning ('Output format "{0}" is unsupported.' -f $f);
            #}
        } # end foreach

    } #end process
} #end function Export-Document

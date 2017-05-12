function OutMD {
<#
    .SYNOPSIS
        MarkDown output plugin for PScribo.
    .DESCRIPTION
        Outputs a markdown file representation of a PScribo document object.
#>
}
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.IO.FileInfo])]
    param (
        ## PScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject] $Document,

        ## Output directory path for the .html file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    begin {

        $pluginName = 'MD';

        <#! OutMD.Internal.ps1 !#>

    } #end begin
    process {

    }

    end{

    }
}
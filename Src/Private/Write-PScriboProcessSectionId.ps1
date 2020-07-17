function Write-PScriboProcessSectionId
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.String] $SectionId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $SectionType,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Length = 40,

        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Indent = 0
    )
    process
    {
        if ($SectionId.Length -gt $Length)
        {
            $sectionDisplayName = '{0}[..]' -f $SectionId.Substring(0, ($Length -4))
        }
        else
        {
            $sectionDisplayName = $SectionId
        }

        $writePScriboMessageParams = @{
            Message = $localized.PluginProcessingSection -f $SectionType, $sectionDisplayName
            Indent  = $Indent
        }
        Write-PScriboMessage @writePScriboMessageParams
    }
}

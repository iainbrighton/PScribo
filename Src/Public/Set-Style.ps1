function Set-Style
{
<#
    .SYNOPSIS
        Sets the style for an individual table row or cell.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
    [OutputType([System.Object])]
    param
    (
        ## PSCustomObject to apply the style to
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Object[]] [Ref] $InputObject,

        ## PScribo style Id to apply
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [System.String] $Style,

        ## Property name(s) to apply the selected style to. Leave blank to apply the style to the entire row.
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]] $Property = '',

        ## Passes the modified object back to the pipeline
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $PassThru
    )
    begin
    {
        if (-not (Test-PScriboStyle -Name $Style))
        {
            Write-Error ($localized.UndefinedStyleError -f $Style)
            return
        }
    }
    process
    {
        foreach ($object in $InputObject)
        {
            foreach ($p in $Property)
            {
                ## If $Property not set, __Style will apply to the whole row.
                $propertyName = '{0}__Style' -f $p
                $object | Add-Member -MemberType NoteProperty -Name $propertyName -Value $Style -Force
            }
        }
        if ($PassThru)
        {
            return $object
        }
    }
}

function ConvertTo-JustifiedString
{
<#
    .SYNOPSIS
        Justifies a block of text using the specified alignment.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Object[]] $InputObject,

        ## Text alignment
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Left','Center','Right','Justify')]
        [System.String] $Align = 'Left',

        ## Text width
        [Parameter()]
        [System.Int32] $Width = ($Host.UI.RawUI.BufferSize.Width -1)
    )
    process
    {
        foreach ($object in $InputObject)
        {
            $objectString = $object.ToString()
            if (($Align -eq 'Left') -or ($objectString.Length -ge $Width))
            {
                $justifiedString = $objectString
            }
            else
            {
                $paddingLength = $Width - $objectString.Length
                if ($Align -eq 'Center')
                {
                    $paddingLength = ($paddingLength /2) -as [System.Int32]
                }
                $padding = ''.PadRight($paddingLength, ' ')
                $justifiedString = '{0}{1}' -f $padding, $objectString
            }
            Write-Output -InputObject $justifiedString
        }
    }
}

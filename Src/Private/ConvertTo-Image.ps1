function ConvertTo-Image
{
<#
    .SYNOPSIS
        Creates an image from a byte[]
#>
    [CmdletBinding()]
    [OutputType([System.Drawing.Image])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Byte[]] $Bytes
    )
    process
    {
        try
        {
            [System.IO.MemoryStream] $memoryStream = New-Object -TypeName 'System.IO.MemoryStream' -ArgumentList @(,$Bytes)
            [System.Drawing.Image] $image = [System.Drawing.Image]::FromStream($memoryStream)
            Write-Output -InputObject $image
        }
        catch
        {
            $_
        }
        finally
        {
            if ($null -ne $memoryStream)
            {
                $memoryStream.Close()
            }
        }
    }
}

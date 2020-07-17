function Get-UriBytes
{
<#
    .SYNOPSIS
        Gets an image's content as a byte[]
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns','')]
    [OutputType([System.Byte[]])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Uri] $Uri
    )
    process
    {
        try
        {
            $webClient = New-Object -TypeName 'System.Net.WebClient'
            [System.IO.Stream] $contentStream = $webClient.OpenRead($uri.AbsoluteUri)
            [System.IO.MemoryStream] $memoryStream = New-Object System.IO.MemoryStream
            $contentStream.CopyTo($memoryStream)
            return $memoryStream.ToArray()
        }
        catch
        {
            $_
        }
        finally
        {
            if ($null -ne $memoryStream) { $memoryStream.Close() }
            if ($null -ne $contentStream) { $contentStream.Close() }
            if ($null -ne $webClient) { $webClient.Dispose() }
        }
    }
}

function Copy-Object
{
<#
    .SYNOPSIS
        Clones a .NET object by serializing and deserializing it.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [System.Object] $InputObject
    )
    process
    {
        try
        {
            $stream = New-Object IO.MemoryStream
            $formatter = New-Object Runtime.Serialization.Formatters.Binary.BinaryFormatter
            $formatter.Serialize($stream, $InputObject)
            $stream.Position = 0
            return $formatter.Deserialize($stream)
        }
        catch
        {
            Write-Error -ErrorRecord $_
        }
        finally
        {
            $stream.Dispose()
        }
    }
}

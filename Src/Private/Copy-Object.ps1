function Copy-Object
{
<#
    .SYNOPSIS
        Clones a .NET object by serializing and deserializing it.
    .NOTES
        PowerShell 7.4 throws a "Type 'System.Management.Automation.PSObject' is not marked as serializable." exception.
        This function has been replaced by 'ConvertTo-Json | ConvertFrom-Json' to serialize and deserialize into a
        "clone" object. This works for PScribo objects but may not work for other nested objects in script blocks.
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

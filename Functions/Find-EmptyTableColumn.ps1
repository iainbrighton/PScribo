function Find-EmptyTableColumn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [PSObject[]]
        $InputObject
    )
    begin {
        if ($null -eq $InputObject) {
            throw 'InputObject cannot be null'
        }
        if ($InputObject -is [string] -or ($InputObject -isnot [System.Object] -and $InputObject -isnot [System.Array])) {
            throw 'InputObject must be an object or array of objects'
        }
    }
}

function Out-JsonTOC {
    <#
    .SYNOPSIS
        Output formatted Table of Contents
#>
    begin {
        ## Initializing TOC object
        $tocBuilder = [ordered]@{}
    }
    process {
        ## Populating TOC
        ## Disregarding section numbering as its highly beneficial when parsing JSON after the fact
        foreach ($tocEntry in $Document.TOC) {
            [ref] $null = $tocBuilder.Add($tocEntry.Number, $tocEntry.Name)
        }

        return ($tocBuilder)
    }
}

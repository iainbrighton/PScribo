function Invoke-PScriboSectionLevel
{
<#
    .SYNOPSIS
        Nested function that processes each document/TOC nested section
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Section,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Number
    )
    process
    {
        if ($pscriboDocument.Options['ForceUppercaseSection'])
        {
            $Section.Name = $Section.Name.ToUpper();
        }

        ## Set this section's level
        $Section.Number = $Number;
        $Section.Level = $Number.Split('.').Count -1;
        ### Add to the TOC
        $tocEntry = [PScustomObject] @{ Id = $Section.Id; Number = $Number; Level = $Section.Level; Name = $Section.Name; }
        [ref] $null = $pscriboDocument.TOC.Add($tocEntry);
        ## Set sub-section level seed
        $minorNumber = 1;
        foreach ($s in $Section.Sections)
        {
            if ($s.Type -like '*.Section' -and -not $s.IsExcluded)
            {
                $sectionNumber = ('{0}.{1}' -f $Number, $minorNumber).TrimStart('.');  ## Calculate section version
                Invoke-PScriboSectionLevel -Section $s -Number $sectionNumber;
                $minorNumber++;
            }
        } #end foreach section
    }
}

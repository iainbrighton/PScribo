function Invoke-PScriboSection
{
<#
    .SYNOPSIS
        Processes the document/TOC section versioning each level, i.e. 1.2.2.3

    .NOTES
        This is an internal function and should not be called directly.
#>
    [CmdletBinding()]
    param ( )
    process
    {
        $majorNumber = 1;
        foreach ($s in $pscriboDocument.Sections)
        {
            if ($s.Type -like '*.Section')
            {
                if ($pscriboDocument.Options['ForceUppercaseSection'])
                {
                    $s.Name = $s.Name.ToUpper();
                }
                if (-not $s.IsExcluded)
                {
                    Invoke-PScriboSectionLevel -Section $s -Number $majorNumber;
                    $majorNumber++;
                }
            }
        }
    }
}

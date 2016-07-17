        #region TOC Private Functions

        function New-PScriboTOC {
        <#
            .SYNOPSIS
                Initializes a new PScribo Table of Contents (TOC) object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(ValueFromPipeline)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name = 'Contents'
            )
            process {

                $typeName = 'PScribo.TOC';
                if ($pscriboDocument.Options['ForceUppercaseSection']) {
                    $Name = $Name.ToUpper();
                }
                $pscriboDocument.Properties['TOCs']++;
                $pscriboTOC = [PSCustomObject] @{
                    Id = [System.Guid]::NewGuid().ToString();
                    Name = $Name;
                    Type = $typeName;
                }
                return $pscriboTOC;

            } #end process
        } #end function New-PScriboLTOC

        #endregion TOC Private Functions

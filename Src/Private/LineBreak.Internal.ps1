        #region LineBreak Private Functions

        function New-PScriboLineBreak {
        <#
            .SYNOPSIS
                Initializes a new PScribo line break object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = [System.Guid]::NewGuid().ToString()
            )
            process {

                $typeName = 'PScribo.LineBreak';
                $pscriboDocument.Properties['LineBreaks']++;
                $pscriboLineBreak = [PSCustomObject] @{
                    Id = $Id;
                    Type = $typeName;
                }
                return $pscriboLineBreak;

            }
        } #end function New-PScriboLineBreak

        #endregion LineBreak Private Functions

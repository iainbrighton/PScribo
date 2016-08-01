        #region BlankLine Private Functions

        function New-PScriboBlankLine {
        <#
            .SYNOPSIS
                Initializes a new PScribo blank line break.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                [Parameter(ValueFromPipeline)]
                [System.UInt32] $Count = 1
            )
            process {

                $typeName = 'PScribo.BlankLine';
                $pscriboDocument.Properties['BlankLines']++;
                $pscriboBlankLine = [PSCustomObject] @{
                    Id = [System.Guid]::NewGuid().ToString();
                    LineCount = $Count;
                    Type = $typeName;
                }
                return $pscriboBlankLine;

            }
        } #end function New-PScriboBlankLine

        #endregion BlankLine Private Functions

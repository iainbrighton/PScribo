        #region Section Private Functions

        function New-PScriboSection {
        <#
            .SYNOPSIS
                Initializes new PScribo section object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PScribo section heading/name.
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,

                ## PScribo style applied to document section.
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.String] $Style = $null,

                ## Section is excluded from TOC/section numbering.
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $IsExcluded
            )
            process {

                $typeName = 'PScribo.Section';
                $pscriboDocument.Properties['Sections']++;
                $pscriboSection = [PSCustomObject] @{
                    Id = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                    Level = 0;
                    Number = '';
                    Name = $Name;
                    Type = $typeName;
                    Style = $Style;
                    IsExcluded = $IsExcluded;
                    Sections = (New-Object -TypeName System.Collections.ArrayList);
                }
                return $pscriboSection;

            } #end process
        } #end function new-pscribosection

        #endregion Section Private Functions

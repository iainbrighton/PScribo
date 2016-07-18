        #region Document Private Functions

        function New-PScriboDocument {
        <#
            .SYNOPSIS
                Initializes a new PScript document object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## PScribo document name
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name
            )
            process {

                WriteLog -Message ($localized.DocumentProcessingStarted -f $Name);
                $typeName = 'PScribo.Document';
                $pscriboDocument = [PSCustomObject] @{
                    Id = $Name.Replace(' ', '').ToUpper();
                    Type = $typeName;
                    Name = $Name;
                    Sections = New-Object -TypeName System.Collections.ArrayList;
                    Options = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Properties = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Styles = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    TableStyles = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    DefaultStyle = $null;
                    DefaultTableStyle = $null;
                    TOC = New-Object -TypeName System.Collections.ArrayList;
                }
                GlobalOption -MarginTopAndBottom 72 -MarginLeftAndRight 54 -PageSize A4 -DefaultFont 'Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif' -Verbose:$false;
                ## Set "default" styles
                Style -Name Normal -Default;
                Style -Name Title -Size 28 -Color 0072af;
                Style -Name TOC -Size 16 -Color 0072af;
                Style -Name 'Heading 1' -Size 16 -Color 0072af;
                Style -Name 'Heading 2' -Size 14 -Color 0072af;
                Style -Name 'Heading 3' -Size 12 -Color 0072af;
                Style -Name TableDefaultHeading -Size 11 -Color fff -Bold -BackgroundColor 4472c4;
                Style -Name TableDefaultRow -Size 11;
                Style -Name TableDefaultAltRow -BackgroundColor d0ddee;
                Style -Name Footer -Size 8 -Color 0072af;
                TableStyle TableDefault -BorderWidth 1 -BorderColor 2a70be -HeaderStyle TableDefaultHeading -RowStyle TableDefaultRow -AlternateRowStyle TableDefaultAltRow -Default;
                return $pscriboDocument;

            } #end process
        } #end function NewPScriboDocument

        function Process-PScriboSection {
        <#
            .SYNOPSIS
                Processes the document/TOC section versioning each level, i.e. 1.2.2.3
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            param ( )

            function Process-PScriboSectionLevel {
            <#
                .SYNOPSIS
                    Nested function that processes each document/TOC nested section
            #>
                [CmdletBinding()]
                param (
                    [Parameter(Mandatory)]
                    [ValidateNotNull()]
                    [PSCustomObject] $Section,

                    [Parameter(Mandatory)]
                    [ValidateNotNullOrEmpty()]
                    [System.String] $Number
                )

                if ($pscriboDocument.Options['ForceUppercaseSection']) {
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
                foreach ($s in $Section.Sections) {
                    if ($s.Type -like '*.Section' -and -not $s.IsExcluded) {
                        $sectionNumber = ('{0}.{1}' -f $Number, $minorNumber).TrimStart('.');  ## Calculate section version
                        Process-PScriboSectionLevel -Section $s -Number $sectionNumber;
                        $minorNumber++;
                    }
                } #end foreach section
            } #end function Process-PScriboSectionLevel

            $majorNumber = 1;
            foreach ($s in $pscriboDocument.Sections) {
                if ($s.Type -like '*.Section') {
                    if ($pscriboDocument.Options['ForceUppercaseSection']) {
                        $s.Name = $s.Name.ToUpper();
                    }
                    if (-not $s.IsExcluded) {
                        Process-PScriboSectionLevel -Section $s -Number $majorNumber;
                        $majorNumber++;
                    }
                } #end if
            } #end foreach
        } #end function process-psscribosection

        #endregion Document Private Functions

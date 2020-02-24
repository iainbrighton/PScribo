        #region Document Private Functions

        function New-PScriboDocument {
        <#
            .SYNOPSIS
                Initializes a new PScript document object.

            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseLiteralInitializerForHashtable','')]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param
            (
                ## PScribo document name
                [Parameter(Mandatory)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,

                ## PScribo document Id
                [Parameter()]
                [ValidateNotNullOrEmpty()]
                [System.String] $Id = $Name.Replace(' ','')
            )
            begin
            {
                if ($(Test-CharsInPath -Path $Name -SkipCheckCharsInFolderPart -Verbose:$false) -eq 3 ) {
                    throw -Message ($localized.IncorrectCharsInName);
                }
            }
            process
            {
                WriteLog -Message ($localized.DocumentProcessingStarted -f $Name);
                $typeName = 'PScribo.Document';
                $pscriboDocument = [PSCustomObject] @{
                    Id                = $Id.ToUpper();
                    Type              = $typeName;
                    Name              = $Name;
                    Sections          = New-Object -TypeName System.Collections.ArrayList;
                    Options           = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Properties        = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    Styles            = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    TableStyles       = New-Object -TypeName System.Collections.Hashtable([System.StringComparer]::InvariantCultureIgnoreCase);
                    DefaultStyle      = $null;
                    DefaultTableStyle = $null;
                    TOC               = New-Object -TypeName System.Collections.ArrayList;
                }
                $defaultDocumentOptionParams = @{
                    MarginTopAndBottom = 72;
                    MarginLeftAndRight = 54;
                    Orientation        = 'Portrait';
                    PageSize           = 'A4';
                    DefaultFont        = 'Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif';
                }
                DocumentOption @defaultDocumentOptionParams -Verbose:$false;

                ## Set "default" styles
                Style -Name Normal -Size 11 -Default -Verbose:$false;
                Style -Name Title -Size 28 -Color '0072af' -Verbose:$false;
                Style -Name TOC -Size 16 -Color '0072af' -Hide -Verbose:$false;
                Style -Name 'Heading 1' -Size 16 -Color '0072af' -Verbose:$false;
                Style -Name 'Heading 2' -Size 14 -Color '0072af' -Verbose:$false;
                Style -Name 'Heading 3' -Size 12 -Color '0072af' -Verbose:$false;
                Style -Name 'Heading 4' -Size 11 -Color '2f5496' -Italic -Verbose:$false;
                Style -Name 'Heading 5' -Size 11 -Color '2f5496' -Verbose:$false;
                Style -Name 'Heading 6' -Size 11 -Color '1f3763' -Verbose:$false;
                Style -Name TableDefaultHeading -Size 11 -Color 'fff' -BackgroundColor '4472c4' -Bold -Verbose:$false;
                Style -Name TableDefaultRow -Size 11 -Verbose:$false;
                Style -Name TableDefaultAltRow -Size 11 -BackgroundColor 'd0ddee' -Verbose:$false;
                Style -Name TableDefaultColumn -Size 11 -Color 'fff' -BackgroundColor '4472c4' -Bold -Verbose:$false;
                Style -Name Footer -Size 8 -Color 0072af -Hide -Verbose:$false;
                $tableDefaultStyleParams = @{
                    Id                = 'TableDefault'
                    BorderWidth       = 1;
                    BorderColor       = '2a70be'
                    HeaderStyle       = 'TableDefaultHeading';
                    RowStyle          = 'TableDefaultRow'
                    AlternateRowStyle = 'TableDefaultAltRow'
                }
                TableStyle @tableDefaultStyleParams -Default -Verbose:$false;
                return $pscriboDocument;

            } #end process
        } #end function NewPScriboDocument

        function Invoke-PScriboSection {
        <#
            .SYNOPSIS
                Processes the document/TOC section versioning each level, i.e. 1.2.2.3

            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            param ( )

            function Invoke-PScriboSectionLevel {
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

            } #end function Invoke-PScriboSectionLevel

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

        function SetIsSectionBreakEnd {
        <#
            Sets the IsSectionBreak end on the last (nested) paragraph/subsection (required by Word plugin)
        #>
            [CmdletBinding()]
            param
            (
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Management.Automation.PSObject] $Section
            )
            process
            {
                $Section.Sections |
                    Where-Object { $_.Type -in 'PScribo.Section','PScribo.Paragraph' } |
                        Select-Object -Last 1 | ForEach-Object {
                            if ($PSItem.Type -eq 'PScribo.Paragraph')
                            {
                                $PSItem.IsSectionBreakEnd = $true
                            }
                            else
                            {
                                SetIsSectionBreakEnd -Section $PSItem
                            }
                        }
            }
        }

        #endregion Document Private Functions

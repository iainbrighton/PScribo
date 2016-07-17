        #region OutXml Private Functions

        function OutXmlSection {
        <#
            .SYNOPSIS
                Output formatted Xml section.
        #>
            [CmdletBinding()]
            param (
                ## PScribo document section
                [Parameter(Mandatory, ValueFromPipeline)]
                [System.Object] $Section
            )
            process {

                $sectionId = ($Section.Id -replace '[^a-z0-9-_\.]','').ToLower();
                $element = $xmlDocument.CreateElement($sectionId);
                [ref] $null = $element.SetAttribute("name", $Section.Name);
                foreach ($s in $Section.Sections.GetEnumerator()) {
                    if ($s.Id.Length -gt 40) { $sectionId = '{0}..' -f $s.Id.Substring(0,38); }
                    else { $sectionId = $s.Id; }
                    $currentIndentationLevel = 1;
                    if ($null -ne $s.PSObject.Properties['Level']) { $currentIndentationLevel = $s.Level +1; }
                    WriteLog -Message ($localized.PluginProcessingSection -f $s.Type, $sectionId) -Indent $currentIndentationLevel;
                    switch ($s.Type) {
                        'PScribo.Section' { [ref] $null = $element.AppendChild((OutXmlSection -Section $s)); }
                        'PScribo.Paragraph' { [ref] $null = $element.AppendChild((OutXmlParagraph -Paragraph $s)); }
                        'PScribo.Table' { [ref] $null = $element.AppendChild((OutXmlTable -Table $s)); }
                        'PScribo.PageBreak' { } ## Page breaks are not implemented for Xml output
                        'PScribo.LineBreak' { } ## Line breaks are not implemented for Xml output
                        'PScribo.BlankLine' { } ## Blank lines are not implemented for Xml output
                        'PScribo.TOC' { } ## TOC is not implemented for Xml output
                        Default {
                            WriteLog -Message ($localized.PluginUnsupportedSection -f $s.Type) -IsWarning;
                        }
                    } #end switch
                } #end foreach
                return $element;

            } #end process
        } #end function outxmlsection


        function OutXmlParagraph {
        <#
            .SYNOPSIS
                Output formatted Xml paragraph.
        #>
            [CmdletBinding()]
            param (
                ## PScribo paragraph object
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Paragraph
            )
            process {

                if (-not ([string]::IsNullOrEmpty($Paragraph.Value))) {
                    ## Value override specified
                    $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
                    $paragraphElement = $xmlDocument.CreateElement($paragraphId);
                    [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Value));
                } #end if
                elseif ([string]::IsNullOrEmpty($Paragraph.Text)) {
                    ## No Id/Name specified, therefore insert as a comment
                    $paragraphElement = $xmlDocument.CreateComment((' {0} ' -f $Paragraph.Id));
                } #end elseif
                else {
                    ## Create an element with the Id/Name
                    $paragraphId = ($Paragraph.Id -replace '[^a-z0-9-_\.]','').ToLower();
                    $paragraphElement = $xmlDocument.CreateElement($paragraphId);
                    [ref] $null = $paragraphElement.AppendChild($xmlDocument.CreateTextNode($Paragraph.Text));
                } #end else
                return $paragraphElement;

            } #end process
        } #end function outxmlparagraph


        function OutXmlTable {
        <#
            .SYNOPSIS
                Output formatted Xml table.
        #>
            [CmdletBinding()]
            param (
                ## PScribo table object
                [Parameter(Mandatory, ValueFromPipeline)]
                [ValidateNotNull()]
                [System.Object] $Table
            )
            process {

                $tableId = ($Table.Id -replace '[^a-z0-9-_\.]','').ToLower();
                $tableElement = $element.AppendChild($xmlDocument.CreateElement($tableId));
                [ref] $null = $tableElement.SetAttribute('name', $Table.Name);
                foreach ($row in $Table.Rows) {
                    $groupElement = $tableElement.AppendChild($xmlDocument.CreateElement('group'));
                    foreach ($property in $row.PSObject.Properties) {
                        if (-not ($property.Name).EndsWith('__Style')) {
                            $propertyId = ($property.Name -replace '[^a-z0-9-_\.]','').ToLower();
                            $rowElement = $groupElement.AppendChild($xmlDocument.CreateElement($propertyId));
                            ## Only add the Name attribute if there's a difference
                            if ($property.Name -ne $propertyId) {
                                [ref] $null = $rowElement.SetAttribute('name', $property.Name);
                            }
                            [ref] $null = $rowElement.AppendChild($xmlDocument.CreateTextNode($row.($property.Name)));
                        } #end if
                    } #end foreach property
                } #end foreach row
                return $tableElement;

            } #end process
        } #end outxmltable

        #endregion OutXml Private Functions

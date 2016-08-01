        #region Paragraph Private Functions

        function New-PScriboParagraph {
        <#
            .SYNOPSIS
                Initializes a new PScribo paragraph object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [OutputType([System.Management.Automation.PSCustomObject])]
            param (
                ## Paragraph Id (and Xml) element name
                [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
                [ValidateNotNullOrEmpty()]
                [System.String] $Name,

                ## Paragraph text. If empty $Name/Id will be used.
                [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
                [AllowNull()]
                [System.String] $Text = $null,

                ## Ouptut value override, i.e. for Xml elements. If empty $Text will be used.
                [Parameter(ValueFromPipelineByPropertyName, Position = 2)]
                [AllowNull()]
                [System.String] $Value = $null,

                ## Paragraph style Name/Id reference.
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.String] $Style = $null,

                ## No new line - ONLY IMPLEMENTED FOR TEXT OUTPUT
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $NoNewLine,

                ## Override the bold style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Bold,

                ## Override the italic style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Italic,

                ## Override the underline style
                [Parameter(ValueFromPipelineByPropertyName)]
                [System.Management.Automation.SwitchParameter] $Underline,

                ## Override the font name(s)
                [Parameter(ValueFromPipelineByPropertyName)]
                [ValidateNotNullOrEmpty()]
                [System.String[]] $Font,

                ## Override the font size (pt)
                [Parameter(ValueFromPipelineByPropertyName)]
                [AllowNull()]
                [System.UInt16] $Size = $null,

                ## Override the font color/colour
                [Parameter(ValueFromPipelineByPropertyName)]
                [Alias('Colour')]
                [AllowNull()]
                [System.String] $Color = $null,

                ## Tab indent
                [Parameter()]
                [ValidateRange(0,10)]
                [System.Int32] $Tabs = 0
            )
            begin {

                if (-not ([string]::IsNullOrEmpty($Text))) {
                    $Name = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                }
                if ($Color) {
                    $Color = Resolve-PScriboStyleColor -Color $Color;
                }

            } #end begin
            process {

                $typeName = 'PScribo.Paragraph';
                $pscriboDocument.Properties['Paragraphs']++;
                $pscriboParagraph = [PSCustomObject] @{
                    Id = $Name;
                    Text = $Text;
                    Type = $typeName;
                    Style = $Style;
                    Value = $Value;
                    NewLine = !$NoNewLine;
                    Tabs = $Tabs;
                    Bold = $Bold;
                    Italic = $Italic;
                    Underline = $Underline;
                    Font = $Font;
                    Size = $Size;
                    Color = $Color;
                }
                return $pscriboParagraph;

            } #end process
        } #end function New-PScriboParagraph

        #endregion Paragraph Private Functions

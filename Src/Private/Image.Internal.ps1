        #region Image Private Functions

        function New-PScriboImage {
        <#
            .SYNOPSIS
                Initializes a new PScribo Image object.
            .NOTES
                This is an internal function and should not be called directly.
        #>
            [CmdletBinding()]
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','')]
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
                ## FilePath
                [Parameter()]
                [System.String] $FilePath,
                ## Tab indent
                [Parameter()]
                [ValidateRange(0,10)]
                [System.Int32] $Tabs = 0
            )
            begin {

                if (-not ([string]::IsNullOrEmpty($Text))) {
                    $Name = $Name.Replace(' ', $pscriboDocument.Options['SpaceSeparator']).ToUpper();
                }

            } #end begin
            process {

                $typeName = 'PScribo.Image';
                $pscriboDocument.Properties['Image']++;
                $pscriboParagraph = [PSCustomObject] @{
                    Id = $Name;
                    Text = $Text;
                    Type = $typeName;
                    Value = $Value;
                    FilePath = $FilePath;
                    Tabs = $Tabs;

                }
                return $pscriboImage;

            } #end process
        } #end function New-PScriboParagraph

        #endregion Paragraph Private Functions
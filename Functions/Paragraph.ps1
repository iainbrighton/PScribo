function Paragraph {
<#
    .SYNOPSIS
        Initializes a new PScribo paragraph object.
#>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCustomObject])]
    param (
        ## Paragraph Id and Xml element name
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Name,

        ## Paragraph text. If empty $Name/Id will be used.
        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [AllowNull()]
        [System.String] $Text = $null,

        ## Output value override, i.e. for Xml elements. If empty $Text will be used.
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
        [AllowNull()]
        [System.String] $Color = $null,

        ## Tab indent
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateRange(0,10)]
        [System.Int32] $Tabs = 0
    )
    begin {

        <#! Paragraph.Internal.ps1 !#>

    } #end begin
    process {

        if ($Name.Length -gt 40) { $paragraphDisplayName = '{0}[..]' -f $Name.Substring(0,36); }
        else { $paragraphDisplayName = $Name; }
        WriteLog -Message ($localized.ProcessingParagraph -f $paragraphDisplayName);
        return (New-PScriboParagraph @PSBoundParameters);

    } #end process
} #end function Paragraph

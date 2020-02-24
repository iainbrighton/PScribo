function OutWordSettingsDocument
{
<#
    .SYNOPSIS
        Outputs Office Open XML settings document
#>
    [CmdletBinding()]
    [OutputType([System.Xml.XmlDocument])]
    param
    (
        [Parameter()]
        [System.Management.Automation.SwitchParameter] $UpdateFields
    )
    process
    {
        ## Create the Style.xml document
        $xmlnsMain = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        # <w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        # xmlns:o="urn:schemas-microsoft-com:office:office"
        # xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        # xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
        # xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"
        # xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        # xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
        # xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
        # xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main"
        # mc:Ignorable="w14 w15">
        $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $settingsDocument.AppendChild($settingsDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $settings = $settingsDocument.AppendChild($settingsDocument.CreateElement('w', 'settings', $xmlnsMain))
        ## Set compatibility mode to Word 2013
        $compat = $settings.AppendChild($settingsDocument.CreateElement('w', 'compat', $xmlnsMain))
        $compatSetting = $compat.AppendChild($settingsDocument.CreateElement('w', 'compatSetting', $xmlnsMain))
        [ref] $null = $compatSetting.SetAttribute('name', $xmlnsMain, 'compatibilityMode')
        [ref] $null = $compatSetting.SetAttribute('uri', $xmlnsMain, 'http://schemas.microsoft.com/office/word')
        [ref] $null = $compatSetting.SetAttribute('val', $xmlnsMain, 15)

        if ($UpdateFields)
        {
            $wupdateFields = $settings.AppendChild($settingsDocument.CreateElement('w', 'updateFields', $xmlnsMain))
            [ref] $null = $wupdateFields.SetAttribute('val', $xmlnsMain, 'true')
        }
        return $settingsDocument
    }
}

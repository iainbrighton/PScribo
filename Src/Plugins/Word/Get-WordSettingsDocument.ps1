function Get-WordSettingsDocument
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
        $xmlns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        $settingsDocument = New-Object -TypeName 'System.Xml.XmlDocument'
        [ref] $null = $settingsDocument.AppendChild($settingsDocument.CreateXmlDeclaration('1.0', 'utf-8', 'yes'))
        $settings = $settingsDocument.AppendChild($settingsDocument.CreateElement('w', 'settings', $xmlns))

        ## Set compatibility mode to Word 2013
        $compat = $settings.AppendChild($settingsDocument.CreateElement('w', 'compat', $xmlns))
        $compatSetting = $compat.AppendChild($settingsDocument.CreateElement('w', 'compatSetting', $xmlns))
        [ref] $null = $compatSetting.SetAttribute('name', $xmlns, 'compatibilityMode')
        [ref] $null = $compatSetting.SetAttribute('uri', $xmlns, 'http://schemas.microsoft.com/office/word')
        [ref] $null = $compatSetting.SetAttribute('val', $xmlns, 15)

        if ($UpdateFields)
        {
            $wupdateFields = $settings.AppendChild($settingsDocument.CreateElement('w', 'updateFields', $xmlns))
            [ref] $null = $wupdateFields.SetAttribute('val', $xmlns, 'true')
        }

        return $settingsDocument
    }
}

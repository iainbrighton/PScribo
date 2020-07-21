@{
    RootModule = 'PScribo.psm1'
    ModuleVersion = '0.9.1'
    GUID = '058eab05-b7bc-4f8b-a2d1-737cc664b12b'
    Author = 'Iain Brighton'
    CompanyName = 'Virtual Engine'
    Copyright = '(c) 2020 Iain Brighton. All rights reserved.'
    Description = 'PScribo documentation Powershell module/framework.'
    PowerShellVersion = '3.0'
    FunctionsToExport = @(
                            'BlankLine',
                            'Document',
                            'DocumentOption',
                            'Export-Document',
                            'Footer',
                            'Header',
                            'Image',
                            'LineBreak',
                            'PageBreak',
                            'Paragraph',
                            'Section',
                            'Set-Style',
                            'Style',
                            'Table',
                            'TableStyle',
                            'Text',
                            'TOC',
                            'Write-PScriboMessage'
                        )
    AliasesToExport   = @(
                            'GlobalOption'
                        )
    PrivateData = @{
        PSData = @{
            Tags = @('Powershell','PScribo','Documentation','Framework','VirtualEngine','Windows','Linux','MacOS','PSEdition_Desktop','PSEdition_Core','Word','Html')
            LicenseUri = 'https://raw.githubusercontent.com/iainbrighton/PScribo/master/LICENSE'
            ProjectUri = 'http://github.com/iainbrighton/PScribo'
            # IconUri = '';
        }
    }
}

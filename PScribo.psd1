@{
    RootModule = 'PScribo.psm1';
    ModuleVersion = '0.7.19';
    GUID = '058eab05-b7bc-4f8b-a2d1-737cc664b12b';
    Author = 'Iain Brighton';
    CompanyName = 'Virtual Engine';
    Copyright = '(c) 2017 Iain Brighton. All rights reserved.';
    Description = 'PScribo documentation Powershell module/framework.';
    PowerShellVersion = '3.0';
    FunctionsToExport = @(
                            'BlankLine',
                            'Document',
                            'DocumentOption',
                            'Export-Document',
                            'LineBreak',
                            'PageBreak',
                            'Paragraph',
                            'Section',
                            'Style',
                            'Set-Style',
                            'Table',
                            'TableStyle',
                            'TOC'
                        );
    AliasesToExport   = @(
                            'GlobalOption'
                        );
    PrivateData = @{
        PSData = @{
            Tags = @('Powershell','PScribo','Documentation','Framework','VirtualEngine')
            LicenseUri = 'https://raw.githubusercontent.com/iainbrighton/PScribo/master/LICENSE';
            ProjectUri = 'http://github.com/iainbrighton/PScribo'
            # IconUri = '';
        } # End of PSData hashtable
    } # End of PrivateData hashtable
}

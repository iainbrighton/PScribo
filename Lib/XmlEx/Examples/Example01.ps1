[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '')]
param()

Import-Module -Name XmlEx -Force;

$x = XmlDocument {
    XmlDeclaration -Encoding 'utf-8'
    XmlNamespace -Uri 'http://www.w3.org/XML/1998/namespace'
    XmlNamespace -Prefix 'v' -Uri 'http://mycustom/namespace'
    XmlNamespace -Prefix 'w' -Uri 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    XmlElement -Name 'document' -Prefix w {
        XmlComment 'My comment'
        XmlElement -Name 'body' {
            XmlAttribute 'att1' 'value1' -Namespace 'http://www.w3.org/XML/1998/namespace'
            XmlAttribute 'att2' 'value2' -Prefix v
            XmlText 'My body value'
        }
    }
} -Verbose

$x | Format-XmlEx

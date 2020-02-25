[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '')]
param()

Import-Module -Name XmlEx -Force;

$x = XmlDocument {
    XmlDeclaration
}

## Appending an XmlElement to an exising XmlDocument
XmlElement -XmlDocument $x -Name 'rootElement' {
    XmlElement 'TextNode' {
        XmlText 'My text node'
    }
} -Verbose

$x | Format-XmlEx

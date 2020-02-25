[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '')]
param()

Import-Module -Name XmlEx -Force;

$x = XmlDocument {
    XmlDeclaration -Encoding 'utf-8' -Standalone 'yes'
    XmlElement 'rootElement'
}

## Appending an XmlElement to an exising root XmlElement. This is required because
## $x.rootElement is coerced into [System.String] (only applicable to the root node)?
XmlElement -XmlElement $x.SelectSingleNode('/rootElement') -Name 'subElement' {
    XmlElement 'TextNode' {
        XmlText 'My text node'
    }
} -Verbose

$x | Format-XmlEx

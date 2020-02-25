[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '')]
param()

Import-Module -Name XmlEx -Force;

$x = [System.Xml.XmlDocument] @'
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:xml="http://www.w3.org/XML/1998/namespace">
  <!--My comment-->
  <w:body xml:att1="value1" />Sub node</w:document>
'@

## Appending an XmlElement to an exising XmlElement
XmlElement -Name 'appended' -XmlElement $x.document {
    XmlElement 'TextNode' {
        XmlText 'My text node'
    }
} -Verbose

## Appending an XmlAttribute to an exising XmlElement
[ref] $null = XmlAttribute -XmlElement $x.document.appended -Name 'myattribute' -Value 'Rubbish!' -verbose

$x | Format-XmlEx

# en-US
ConvertFrom-StringData @'
    CreatingDocument                     = Creating XmlEx document.
    SettingDocumentNamespace             = Setting document namespace to '{0}'.
    SettingDocumentNamespaceWithPrefix   = Setting document namespace '{0}' with prefix '{1}'.
    SettingDefaultDocumentNamespace      = Setting default document namespace '{0}'.
    AddingElement                        = Adding element '{0}' to element '{1}'.
    AddingAttribute                      = Adding attribute '{0}' to element '{1}'.
    AddingTextNode                       = Adding text node to element '{0}'.
    AddingComment                        = Adding comment to element '{0}'.
    AddingDocumentNamespace              = Adding document namespace '{0}'.
    FinalizingDocument                   = Finalizing XmlEx document.

    ProcessingDocument                   = Process Xml document '{0}'.
    AppendingXmlElementPath              = Appending Xml element/path '{0}'.
    AppendingXmlAttributePath            = Appending attribute '{0}'.
    AppendingXmlTextNodePath             = Appending text node '{0}'.
    SavingDocument                       = Saving Xml document '{0}'.

    XmlExDocumentNotFoundError           = XmlEx document not found/reference not set.
    XmlExElementNotFoundError            = XmlEx element not found/reference not set.
    XmlExNamespaceMissingXmlElementError = Cannot add a 'XmlNamespace' to a document without a root element.
    XmlExDocumentMissingXmlElementError  = Cannot add a 'XmlAttribute' to a document without a root element.
    XmlExInvalidCallOutsideScopeError    = You cannot call '{0}' outside the '{1}' scope.
    XmlExInvalidCallWithinScopeError     = You cannot call '{0}' from within the '{1}' scope. '{0}' must be nested within a '{2}' scope.
    CannotFindPathError                  = Cannot find path '{0}' because it does not exist.

    ShouldProcessOperationWarning        = Performing operation "{0}" on Target "{1}".
    ShouldProcessWarning                 = Continue with this operation?
'@;

# en-US
ConvertFrom-StringData @'
ImportingFile                    = Importing file '{0}'.
InvalidDirectoryPathError        = Path '{0}' is not a valid directory path.'
NoScriptBlockProvidedError       = No PScribo section script block is provided (have you put the open curly brace on the next line?).
InvalidHtmlColorError            = Invalid Html color '{0}' specified.
InvalidHtmlBackgroundColorError  = Invalid Html background color '{0}' specified.
UndefinedTableHeaderStyleError   = Undefined table header style '{0}' specified.
UndefinedTableRowStyleError      = Undefined table row style '{0}' specified.
UndefinedAltTableRowStyleError   = Undefined table alternating row style '{0}' specified.
InvalidTableBorderColorError     = Invalid table border color '{0}' specified.
UndefinedStyleError              = Undefined style '{0}' specified.
OpenPackageError                 = Error opening package '{0}'. Ensure the file in not in use by another process.
IncorrectCharsInPathError        = The incorrect char found in the Path.
HeaderFooterDocumentRootError    = The 'Header' and 'Footer' keywords can only be defined in the document root section.
ParagraphRunRootError            = The 'Text' keyword can only be defined within a 'Paragraph' section.
ListRootError                    = The 'List' keyword can only be defined within a 'Paragraph' or 'Section' block.
ItemRootError                    = The 'Item' keyword can only be defined within a 'List' section.
InvalidCustomNumberStyleError    = The custom number style '{0}' is invalid and must contain at least one '%'.

MaxHeadingLevelWarning           = Html5 supports a maximum of 6 heading levels. Reduce the number of nested Document sections to remove the unsupported tags in the resulting Html output.
TableHeadersWithNoColumnsWarning = Table headers have been specified with no table columns/properties. Headers will be ignored.
TableHeadersCountMismatchWarning = The number of table headers specified does not match the number of specified columns/properties. Headers will be ignored.
ListTableColumnCountWarning      = Table columns widths in list format must be 2. Column widths will be ignored.
TableColumnWidthMismatchWarning  = The specified number of table columns and column widths do not match. Column widths will be ignored.
TableColumnWidthSumWarning       = The table column widths total '{0}'%. Total column width must equal 100%. Column widths will be ignored.
TableWidthOverflowWarning        = The table width overflows the page margin and has been adjusted to '{0}mm'.
ImageHeightPercentageError       = The image height with '-AsPercent' cannot be less-than or equal to 0% and/or greater than 100%.
ImageWidthPercentageError        = The image width with '-AsPercent' cannot be less-than or equal to 0% and/or greater than 100%.
UnexpectedObjectWarning          = Unexpected/unsupported object in document/section '{0}'.
UnexpectedObjectTypeWarning      = Unexpected '{0}' object in document/section '{1}'.
UnsupportedPScriboTypeWarning    = PScribo type '{0}' is not supported in document/section '{1}'.
CannotSetOrientationWarning      = Orientation can only be set on a document root 'Section'. Section orientation will be ignored.
ListTableCaptionRemovedWarning   = List table captions are only supported on tables with a single row. Removing caption from table '{0}'.
FirstPageHeaderOverwriteWarning  = Existing first page header definition overwritten.
DefaultHeaderOverwriteWarning    = Existing default page header definition overwritten.
FirstPageFooterOverwriteWarning  = Existing first page footer definition overwritten.
DefaultFooterOverwriteWarning    = Existing default page footer definition overwritten.
NoNewLineDeprecatedWarning       = The '-NoNewLine' functionality has been deprecated. Use Paragraph runs (Text) to implement this functionality for all output formats.
ValueParameterRemovedWarning     = The 'Paragraph -Value' functionality has been removed and is no longer implemented.
NoPriorListItemWarning           = No 'Item' defined before nested 'List'; nested list will be ignored.

DocumentProcessingStarted        = Document '{0}' processing started.
DocumentInvokePlugin             = Invoking '{0}' plugin.
DocumentExportPluginComplete     = Plugin '{0}' complete.
DocumentOptions                  = Setting global document options.
DocumentOptionSpaceSeparator     = Setting default space separator to '{0}'.
DocumentOptionUppercaseHeadings  = Enabling uppercase headings.
DocumentOptionUppercaseSections  = Enabling uppercase sections.
DocumentOptionSectionNumbering   = Enabling section/heading numbering.
DocumentOptionPageTopMargin      = Setting page top margin to '{0}'mm.
DocumentOptionPageRightMargin    = Setting page right margin to '{0}'mm.
DocumentOptionPageBottomMargin   = Setting page bottom margin to '{0}'mm.
DocumentOptionPageLeftMargin     = Setting page left margin to '{0}'mm.
DocumentOptionPageSize           = Setting page size to '{0}'.
DocumentOptionPageOrientation    = Setting page orientation to '{0}'.
DocumentOptionPageHeight         = Setting page height to '{0}'mm.
DocumentOptionPageWidth          = Setting page width to '{0}'mm.
DocumentOptionDefaultFont        = Setting default font(s) to '{0}'.
ProcessingBlankLine              = Processing blank line.
ProcessingImage                  = Processing image '{0}'.
ProcessingLineBreak              = Processing line break.
ProcessingPageBreak              = Processing page break.
ProcessingParagraph              = Processing paragraph '{0}'.
ProcessingSection                = Processing section '{0}'.
ProcessingSectionStarted         = Processing section '{0}' started.
ProcessingSectionCompleted       = Processing section '{0}' completed.
PluginProcessingSection          = Processing {0} '{1}'.
ProcessingStyle                  = Setting document style '{0}'.
ProcessingTable                  = Processing table '{0}'.
ProcessingTableStyle             = Setting table style '{0}'.
ProcessingTOC                    = Processing table of contents '{0}'.
ProcessingDocumentPart           = Processing document part '{0}'.
WritingDocumentPart              = Writing document part '{0}'.
GeneratingPackageRelationships   = Generating package relationships.
PluginUnsupportedSection         = Unsupported section '{0}'.
DocumentProcessingCompleted      = Document '{0}' processing completed.
TotalProcessingTimeSeconds       = Total processing time '{0:N2}' seconds.
TotalProcessingTimeMinutes       = Total processing time '{0:N2}' minutes.
SavingFile                       = Saving file '{0}'.
ProcessingHeaderStarted          = Processing document header started.
ProcessingHeaderCompleted        = Processing document header completed.
ProcessingFooterStarted          = Processing document footer started.
ProcessingFooterCompleted        = Processing document footer completed.
ProcessingParagraphRunsStarted   = Processing paragraph run(s) started.
ProcessingParagraphRunsCompleted = Processing paragraph run(s) completed.
ProcessingParagraphRun           = Processing paragraph run '{0}'.
ProcessingList                   = Processing list '{0}'.
ProcessingNumberStyle            = Setting number style '{0}'.

# Enhanced Table Messages
ProcessingEnhancedTable         = Processing enhanced table '{0}'.
ProcessingEmptyColumns          = Checking for empty columns in table '{0}'.
ProcessingColumnWidths          = Calculating column widths for table '{0}'.
EmptyColumnsFound               = Found {0} empty columns in table '{1}': {2}.
NoEmptyColumnsFound             = No empty columns found in table '{0}'.
ColumnWidthsCalculated          = Column widths calculated for table '{0}': {1}.
EnhancedTableCompleted          = Enhanced table '{0}' processing completed.

NoInputProvided                 = No input objects provided for table '{0}'.
NoPropertiesFound               = No properties found in input objects for table '{0}'.
'@;

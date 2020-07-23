# Change Log #

## Versions ##

### Unreleased ###

* __BREAKING CHANGE__ - XML output to be deprecated in a future release (#102)
* Adds inline paragraph styling runs - see Example38.ps1 (#75, #97)
  * __BREAKING CHANGE__ - XML paragraph output formatting changed
* Adds `about_PScriboExamples` help file (#103)
  * Run `Get-Help about_PScriboExamples` to return example documentation
* Adds dynamic `-Format` parameter to `Export-Document`
* Adds `about_TextPlugin`, `about_HtmlPlugin` and `about_WordPlugin` help topics

### 0.9.1 ###

* Fixes Word table width bug with landscape orientation

### 0.9.0 ###

* Adds Header and Footer support (#80)
* Corrects table caption numbering (#96)
* Fixes new PSScriptAnalyzer rule and style violations
* Adds LibreOffice compatibility (#99)

### 0.8.0 ###

* Fixes text line ending output and tests on Linux and MacOS
* Adds keyed list support (#87) - see Example33.ps1
* Improves text alignment/justification output
* Adds table caption support (#88) -see Example34.ps1
* Improves Html page rendering - now displays a minimum page size
* Fixes .NET "Unable to determine the identity of the domain" error (https://github.com/AsBuiltReport/AsBuiltReport.Core/issues/17)
* Fixes Word table style header alignment

### 0.7.26 ###

* Fixes bug in Html table header CSS style output
* Fixes unit tests on PowerShell Core
* Adds `TableStyle -Padding` shortcut

### 0.7.25 ###

* Replaces Html `em` output with `rem` to maintain correct font sizing
* Adds page orientation support to document root `Section` s (#81)
* Adds `Image` support (#12)
* Exposes `Write-PScriboMessage` function externally (#89)
* Adds AppVeyor build
* Adds linting tests

### 0.7.24 ###

* Fixes "Non-negative number required" errors in text TOC output (#77)

### 0.7.23 ###

* Fixes Html TOC section hyperlinks with duplicate names (#74)
* Adds warning to errant Document pipeline output
* Adds tab/indentation support to `Section` headers (#73)
* Removes trailing spaces from text table output (#67)

### 0.7.22 ###

* Fixes $null Html output and [String]::Empty Word output (#72)

### 0.7.21 ###

* Fixes custom table style default border color output (#71)
* Fixes bundle creation after internal directory restructure
* Fixes XML output appending rather than overwriting existing files

### 0.7.20 ###

* Fixes unit tests for Pester v4.x compatibility
* Fixes [DateTime] discrepancy in Word and Html table output (#68)
* Fixes errors importing module on non-English locales (#63)
* Corrects PowerShell Core `IsMacOS` environment variable rename

### 0.7.19 ###

* Adds default styles for heading levels 4 - 6 (#59)
* Fixes $PSVersionTable.PSEdition strict mode detection errors on PS3 and PS4
* Fixes bug with nested HTML section output depth warning (#57)

### 0.7.18 ###

* Adds ClassId parameter to TOC for HTML output

### 0.7.17 ###

* Renames 'Hide' parameter to 'Hidden' on Style keyword
  * Adds 'Hide' alias for backwards compatibility
* Fixes errors importing localized data on en-US systems (#49)

### 0.7.16 ###

* Supports hiding styles, e.g. in MS Word (#32)
* Fixes Html and Word decimal localisation issues (#6, #42)
* Adds support for CRLFs in paragraphs and table cells (#46)

### 0.7.15 ###

* Adds Core PowerShell (v6.0.0) support
* Fixes tests in Pester v4
* Adds `Merge-PScriboPluginOptions` method to merge document, plugin and runtime options (#24)
* Adds `-PassThru` option to `Export-Document` (#25)
* Renames `GlobalOption` to `DocumentOption` (#18)
  * Adds `GlobalOption` alias for backwards compatibility

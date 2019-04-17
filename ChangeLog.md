# Change Log #

## Versions ##

### Unreleased ###

* Adds AppVeyor build

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

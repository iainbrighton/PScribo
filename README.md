## PScribo ##
PScribo is an open source project that implements a documentation domain-specific
language (DSL) for Windows Powershell. The latest version is available at
https://github.com/iainbrighton/PScribo.
    
PScribo provides a set of functions that make it easy to create a document-like
structure within Powershell scripts, without having to be concerned with
handling output formatting or supporting multiple output formats. For more
detailed infomation on the documentation DSL, see about_Document.

Pscribo can export documentation in a variety of formats and currently
supports creation of text, xml, html and Microsoft Word formats. 
Additional "plugins" can be created to support future formats if required. For
more detailed information on creating a "plugin" see about_Plugins.

PScribo is available as a Powershell module or a bundle. The bundle release
permits dot-sourcing the PScribo bundle file into an existing Powershell script
or can be placed in its entirety at the beginning of a .ps1 file.

Requires __Powershell 3.0__ or later.

If you find it useful, unearth any bugs or have any suggestions for improvements,
feel free to add an [issue](https://github.com/iainbrighton/PScribo/issues) or
place a comment at the project home page</a>.

##### Installation

* Automatic (via PowerShell v5 or later) __NOT YET PUBLISHED__:
 * Run 'Install-Module PScribo'.
 * Run 'Import-Module PScribo'.
* Manual:
 * Download the [latest release .zip](https://github.com/iainbrighton/PScribo/releases/latest)
 * Extract the .zip into your $PSModulePath, e.g. ~\Documents\WindowsPowerShell\Modules\.
 * Run 'Import-Module PScribo'.

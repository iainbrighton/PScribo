function Out-HtmlDocument
{
<#
    .SYNOPSIS
        Html output plugin for PScribo.

    .DESCRIPTION
        Outputs a Html file representation of a PScribo document object.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','pluginName')]
    [OutputType([System.IO.FileInfo])]
    param
    (
        ## PScribo document object to convert to a text document
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $Document,

        ## Output directory path for the .html file
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String] $Path,

        ### Hashtable of all plugin supported options
        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowNull()]
        [System.Collections.Hashtable] $Options
    )
    process
    {
        $pluginName = 'Html'
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        Write-PScriboMessage -Message ($localized.DocumentProcessingStarted -f $Document.Name)

        ## Merge the document, plugin default and specified/specific plugin options
        $mergePScriboPluginOptionParams = @{
            DefaultPluginOptions = New-PScriboHtmlOption
            DocumentOptions = $Document.Options
            PluginOptions = $Options
        }
        $options = Merge-PScriboPluginOption @mergePScriboPluginOptionParams
        $noPageLayoutStyle = $Options['NoPageLayoutStyle']
        $topMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $options['MarginTop']) -Format 'f2'
        $leftMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $options['MarginLeft']) -Format 'f2'
        $bottomMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $options['MarginBottom']) -Format 'f2'
        $rightMargin = ConvertTo-InvariantCultureString -Object (ConvertTo-Em -Millimeter $options['MarginRight']) -Format 'f2'
        $script:currentPageNumber = 1

        [System.Text.StringBuilder] $htmlBuilder = New-Object System.Text.StringBuilder
        [ref] $null = $htmlBuilder.AppendLine('<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">')
        [ref] $null = $htmlBuilder.AppendLine('<html xmlns="http://www.w3.org/1999/xhtml">')
        [ref] $null = $htmlBuilder.AppendLine('<head><title>{0}</title>' -f $Document.Name)
        [ref] $null = $htmlBuilder.AppendLine('{0}</head><body>' -f (Out-HtmlStyle -Styles $Document.Styles -TableStyles $Document.TableStyles -NoPageLayoutStyle:$noPageLayoutStyle))
        [ref] $null = $htmlBuilder.AppendFormat('<div class="{0}">', $options['PageOrientation'].ToLower())
        [ref] $null = $htmlBuilder.AppendFormat('<div class="{0}" style="padding-top: {1}rem; padding-left: {2}rem; padding-bottom: {3}rem; padding-right: {4}rem;">', $Document.DefaultStyle, $topMargin, $leftMargin, $bottomMargin, $rightMargin).AppendLine()

        [ref] $null = $htmlBuilder.AppendLine((Out-HtmlHeaderFooter -Header -FirstPage))

        $canvasHeight = Get-HtmlCanvasHeight -Orientation $options['PageOrientation']
        [ref] $null = $htmlBuilder.AppendFormat('<div style="min-height: {0}mm" >', $canvasHeight)

        foreach ($subSection in $Document.Sections.GetEnumerator())
        {
            $currentIndentationLevel = 1
            if ($null -ne $subSection.PSObject.Properties['Level'])
            {
                $currentIndentationLevel = $subSection.Level +1
            }
            Write-PScriboProcessSectionId -SectionId $subSection.Id -SectionType $subSection.Type -Indent $currentIndentationLevel

            switch ($subSection.Type)
            {
                'PScribo.Section'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlSection -Section $subSection))
                }
                'PScribo.Paragraph'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlParagraph -Paragraph $subSection))
                }
                'PScribo.Table'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlTable -Table $subSection))
                }
                'PScribo.LineBreak'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlLineBreak))
                }
                'PScribo.PageBreak'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlPageBreak -Orientation $options['PageOrientation']))
                }
                'PScribo.TOC'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlTOC -TOC $subSection))
                }
                'PScribo.BlankLine'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlBlankLine -BlankLine $subSection))
                }
                'PScribo.Image'
                {
                    [ref] $null = $htmlBuilder.Append((Out-HtmlImage -Image $subSection))
                }
                Default
                {
                    Write-PScriboMessage -Message ($localized.PluginUnsupportedSection -f $subSection.Type) -IsWarning
                }
            }
        } #end foreach section

        [ref] $null = $htmlBuilder.AppendLine('</div>') ## Canvas
        [ref] $null = $htmlBuilder.Append((Out-HtmlHeaderFooter -Footer))

        [ref] $null = $htmlBuilder.AppendLine('</div></div></body>')
        $stopwatch.Stop()
        Write-PScriboMessage -Message ($localized.DocumentProcessingCompleted -f $Document.Name)
        $destinationPath = Join-Path $Path ('{0}.html' -f $Document.Name)
        Write-PScriboMessage -Message ($localized.SavingFile -f $destinationPath)
        $htmlBuilder.ToString().TrimEnd() | Out-File -FilePath $destinationPath -Force -Encoding utf8
        [ref] $null = $htmlBuilder

        if ($stopwatch.Elapsed.TotalSeconds -gt 90)
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeMinutes -f $stopwatch.Elapsed.TotalMinutes)
        }
        else
        {
            Write-PScriboMessage -Message ($localized.TotalProcessingTimeSeconds -f $stopwatch.Elapsed.TotalSeconds)
        }

        Write-Output (Get-Item -Path $destinationPath)
    }
}

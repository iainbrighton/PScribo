function Out-HtmlHeaderFooter
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'DefaultHeader')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageHeader')]
        [System.Management.Automation.SwitchParameter] $Header,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'DefaultFooter')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageFooter')]
        [System.Management.Automation.SwitchParameter] $Footer,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageHeader')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FirstPageFooter')]
        [System.Management.Automation.SwitchParameter] $FirstPage
    )
    process
    {
        $headerFooter = Get-PScriboHeaderFooter @PSBoundParameters
        if ($null -ne $headerFooter)
        {
            [System.Text.StringBuilder] $hfBuilder = New-Object System.Text.StringBuilder
            [ref] $null = $hfBuilder.Append('<div>')

            foreach ($subSection in $headerFooter.Sections.GetEnumerator())
            {
                switch ($subSection.Type)
                {
                    'PScribo.Paragraph'
                    {
                        [ref] $null = $hfBuilder.Append((Out-HtmlParagraph -Paragraph $subSection))
                    }
                    'PScribo.Table'
                    {
                        [ref] $null = $hfBuilder.Append((Out-HtmlTable -Table $subSection))
                    }
                    'PScribo.BlankLine'
                    {
                        [ref] $null = $hfBuilder.Append((Out-HtmlBlankLine -BlankLine $subSection))
                    }
                }
            }

            [ref] $null = $hfBuilder.Append('</div>')
            return $hfBuilder.ToString()
        }
    }
}

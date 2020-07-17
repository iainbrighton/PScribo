function Out-TextHeaderFooter
{
<#
    .SYNOPSIS
        Output formatted header/footer.
#>
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

            foreach ($subSection in $headerFooter.Sections.GetEnumerator())
            {
                ## When replacing tokens (by reference), the tokens are removed
                $cloneSubSection = Copy-Object -InputObject $subSection
                switch ($cloneSubSection.Type)
                {
                    'PScribo.Paragraph'
                    {
                        $paragraph = Out-TextParagraph -Paragraph $cloneSubSection
                        [ref] $null = $hfBuilder.Append($paragraph)
                    }
                    'PScribo.Table'
                    {
                        if ($Footer)
                        {
                            ## Add a space before a footer table
                            [ref] $null = $hfBuilder.AppendLine()
                        }

                        $table = Out-TextTable -Table $cloneSubSection
                        [ref] $null = $hfBuilder.Append($table)

                        if ($Header)
                        {
                            ## Add a space after a header table
                            [ref] $null = $hfBuilder.AppendLine()
                        }
                    }
                    'PScribo.BlankLine'
                    {
                        $blankLine = Out-TextBlankLine -BlankLine $subSection
                        [ref] $null = $hfBuilder.Append($blankLine)
                    }
                }
            }

            return $hfBuilder.ToString()
        }
    }
}

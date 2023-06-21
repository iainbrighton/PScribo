function Out-JsonHeaderFooter
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
    begin
    {
        ## Initializing header/footer object
        $hfBuilder = [ordered]@{}

        ## Initializing paragraph counter
        [int]$paragraph = 1

        ## Initializing table counter
        [int]$table = 1
    }
    process
    {
        $headerFooter = Get-PScriboHeaderFooter @PSBoundParameters
        if ($null -ne $headerFooter)
        {
            foreach ($subSection in $headerFooter.Sections.GetEnumerator())
            {
                ## When replacing tokens (by reference), the tokens are removed
                $cloneSubSection = Copy-Object -InputObject $subSection
                switch ($cloneSubSection.Type)
                {
                    'PScribo.Paragraph'
                    {
                        [ref] $null = $hfBuilder.Add("paragraph$($paragraph)", (Out-JsonParagraph -Paragraph $cloneSubSection))
                        $paragraph++
                    }
                    'PScribo.Table'
                    {
                        [ref] $null = $hfBuilder.Add("table$($table)", (Out-JsonTable -Table $cloneSubSection))
                        $table++
                    }
                }
            }

            return $hfBuilder
        }
        else {
            return $null
        }
    }
}

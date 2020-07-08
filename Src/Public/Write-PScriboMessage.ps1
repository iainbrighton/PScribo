function Write-PScriboMessage {
<#
    .SYNOPSIS
        Writes PScribo-formatted message to the verbose, warning or debug streams. Output is
        prefixed with the time and PScribo plugin name.
#>
    [CmdletBinding(DefaultParameterSetName = 'Verbose')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','IsWarning')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','IsDebug')]
    param
    (
        ## Message to send to the stream
        [Parameter(Position = 0, ValueFromPipeline, ParameterSetName = 'Verbose')]
        [Parameter(Position = 0, ValueFromPipeline, ParameterSetName = 'Warning')]
        [Parameter(Position = 0, ValueFromPipeline, ParameterSetName = 'Debug')]
        [ValidateNotNullOrEmpty()]
        [System.String] $Message,

        ## PScribo plugin name
        [Parameter(Position = 1, ValueFromPipelineByPropertyName)]
        [System.String] $Plugin,

        ## Redirect message to the Warning stream
        [Parameter(ParameterSetName = 'Warning')]
        [System.Management.Automation.SwitchParameter] $IsWarning,

        ## Redirect message to the Debug stream
        [Parameter(ParameterSetName = 'Debug')]
        [System.Management.Automation.SwitchParameter] $IsDebug,

        ## Padding/indent section level
        [Parameter(ValueFromPipeline, ParameterSetName = 'Verbose')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Warning')]
        [Parameter(ValueFromPipeline, ParameterSetName = 'Debug')]
        [ValidateNotNullOrEmpty()]
        [System.UInt16] $Indent
    )
    process
    {
        if ([System.String]::IsNullOrEmpty($Plugin))
        {
            ## Attempt to resolve the plugin name from the parent scope
            if (Test-Path -Path Variable:\pluginName)
            {
                $Plugin = Get-Variable -Name pluginName -ValueOnly
            }
            else
            {
                $Plugin = 'Unknown'
            }
        }
        ## Center plugin name
        $pluginPaddingSize = [System.Math]::Floor((10 - $Plugin.Length) / 2)
        $pluginPaddingString = ''.PadRight($pluginPaddingSize)
        $Plugin = '{0}{1}' -f $pluginPaddingString, $Plugin
        $Plugin = $Plugin.PadRight(10)
        $date = Get-Date
        $sectionLevelPadding = ''.PadRight($Indent)
        $formattedMessage = '[ {0} ] [{1}] - {2}{3}' -f $date.ToString('HH:mm:ss:fff'), $Plugin, $sectionLevelPadding, $Message
        switch ($PSCmdlet.ParameterSetName)
        {
            'Warning' { Write-Warning -Message $formattedMessage }
            'Debug' { Write-Debug -Message $formattedMessage }
            Default { Write-Verbose -Message $formattedMessage }
        }
    }
}

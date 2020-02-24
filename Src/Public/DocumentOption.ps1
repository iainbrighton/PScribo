function DocumentOption {
<#
    .SYNOPSIS
        Initializes a new PScribo global/document options/settings.

    .NOTES
        Options are reset upon each invocation.
#>
    [CmdletBinding(DefaultParameterSetName = 'Margin')]
    [Alias('GlobalOption')]
    param
    (
        ## Forces document header to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseHeader,

        ## Forces all section headers to be displayed in upper case.
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $ForceUppercaseSection,

        ## Enable section/heading numbering
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Management.Automation.SwitchParameter] $EnableSectionNumbering,

        ## Default space replacement separator
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Separator')]
        [AllowNull()]
        [ValidateLength(0,1)]
        [System.String] $SpaceSeparator,

        ## Default page top, bottom, left and right margin (pt)
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Margin')]
        [System.UInt16] $Margin = 72,

        ## Default page top and bottom margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginTopAndBottom,

        ## Default page left and right margins (pt)
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'CustomMargin')]
        [System.UInt16] $MarginLeftAndRight,

        ## Default page size
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('A4','Legal','Letter')]
        [System.String] $PageSize = 'A4',

        ## Page orientation
        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Portrait','Landscape')]
        [System.String] $Orientation = 'Portrait',

        ## Default document font(s)
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String[]] $DefaultFont = @('Calibri','Candara','Segoe','Segoe UI','Optima','Arial','Sans-Serif')
    )
    process
    {
        $localized.DocumentOptions | WriteLog;
        if ($SpaceSeparator)
        {
            WriteLog -Message ($localized.DocumentOptionSpaceSeparator -f $SpaceSeparator);
            $pscriboDocument.Options['SpaceSeparator'] = $SpaceSeparator;
        }

        if ($ForceUppercaseHeader)
        {
            $localized.DocumentOptionUppercaseHeadings | WriteLog;
            $pscriboDocument.Options['ForceUppercaseHeader'] = $true;
            $pscriboDocument.Name = $pscriboDocument.Name.ToUpper();
        }

        if ($ForceUppercaseSection)
        {
            $localized.DocumentOptionUppercaseSections | WriteLog;
            $pscriboDocument.Options['ForceUppercaseSection'] = $true;
        }

        if ($EnableSectionNumbering)
        {
            $localized.DocumentOptionSectionNumbering | WriteLog;
            $pscriboDocument.Options['EnableSectionNumbering'] = $true;
        }

        if ($DefaultFont)
        {
            WriteLog -Message ($localized.DocumentOptionDefaultFont -f ([System.String]::Join(', ', $DefaultFont)));
            $pscriboDocument.Options['DefaultFont'] = $DefaultFont;
        }

        if ($PSCmdlet.ParameterSetName -eq 'CustomMargin')
        {
            if ($MarginTopAndBottom -eq 0) { $MarginTopAndBottom = 72; }
            if ($MarginLeftAndRight -eq 0) { $MarginTopAndBottom = 72; }
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $MarginTopAndBottom;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = ConvertPtToMm -Point $MarginLeftAndRight;
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginLeft'];
        }
        else
        {
            $pscriboDocument.Options['MarginTop'] = ConvertPtToMm -Point $Margin;
            $pscriboDocument.Options['MarginBottom'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginLeft'] = $pscriboDocument.Options['MarginTop'];
            $pscriboDocument.Options['MarginRight'] = $pscriboDocument.Options['MarginTop'];
        }
        WriteLog -Message ($localized.DocumentOptionPageTopMargin -f $pscriboDocument.Options['MarginTop']);
        WriteLog -Message ($localized.DocumentOptionPageRightMargin -f $pscriboDocument.Options['MarginRight']);
        WriteLog -Message ($localized.DocumentOptionPageBottomMargin -f $pscriboDocument.Options['MarginBottom']);
        WriteLog -Message ($localized.DocumentOptionPageLeftMargin -f $pscriboDocument.Options['MarginLeft']);

        ## Convert page size
        ($localized.DocumentOptionPageSize -f $PageSize) | WriteLog;
        switch ($PageSize)
        {
            'A4' {
                $pscriboDocument.Options['PageWidth'] = 210.0;
                $pscriboDocument.Options['PageHeight'] = 297.0;
            }
            'Legal' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 355.6;
            }
            'Letter' {
                $pscriboDocument.Options['PageWidth'] = 215.9;
                $pscriboDocument.Options['PageHeight'] = 279.4;
            }
        }

        ## Convert page size
        ($localized.DocumentOptionPageOrientation -f $Orientation) | WriteLog;
        $pscriboDocument.Options['PageOrientation'] = $Orientation;
        $script:currentOrientation = $Orientation;
        ($localized.DocumentOptionPageHeight -f $pscriboDocument.Options['PageHeight']) | WriteLog;
        ($localized.DocumentOptionPageWidth -f $pscriboDocument.Options['PageWidth']) | WriteLog;
    }
}

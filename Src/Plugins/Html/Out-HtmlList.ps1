function Out-HtmlList
{
<#
    .SYNOPSIS
        Output formatted Html list.
#>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','Options')]
    param
    (
        ## List to output
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.PSObject] $List,

        ## List indent
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.Int32] $Indent,

        ## Number style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $NumberStyle,

        ## Bullet style
        [Parameter(ValueFromPipelineByPropertyName)]
        [System.String] $BulletStyle
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboHtmlOption
        }
    }
    process
    {
        $leader = ''.PadRight($Indent, ' ')
        $listBuilder = New-Object -TypeName System.Text.StringBuilder

        if ($List.IsNumbered)
        {
            if ($List.HasNumberStyle)
            {
                $NumberStyle = $List.NumberStyle
            }
            elseif (-not $PSBoundParameters.ContainsKey('NumberStyle'))
            {
                $NumberStyle = $Document.DefaultNumberStyle
            }
            $style = $Document.NumberStyles[$NumberStyle]

            $outHtmlListParams = @{
                NumberStyle = $NumberStyle
            }

            switch ($style.Format)
            {
                'Number'
                {
                    $inlineStyle = 'decimal'
                }
                'Letter'
                {
                    if ($style.Uppercase)
                    {
                        $inlineStyle = 'upper-alpha'
                    }
                    else
                    {
                        $inlineStyle = 'lower-alpha'
                    }
                }
                'Roman'
                {
                    if ($style.Uppercase)
                    {
                        $inlineStyle = 'upper-roman'
                    }
                    else
                    {
                        $inlineStyle = 'lower-roman'
                    }
                }
                'Custom'
                {
                    $inlineStyle = 'decimal'
                }
            }

            if ($List.HasStyle)
            {
                [ref] $null = $listBuilder.AppendFormat('{0}<ol class="{1}" style="list-style-type:{2};">', $leader, $List.Style, $inlineStyle)
            }
            else
            {
                [ref] $null = $listBuilder.AppendFormat('{0}<ol style="list-style-type:{1};">', $leader, $inlineStyle)
            }
            [ref] $null = $listBuilder.AppendLine()
        }
        else
        {
            if ($List.HasBulletStyle)
            {
                $BulletStyle = $List.BulletStyle
            }

            switch ($BulletStyle)
            {
                Circle
                {
                    $inlineStyle = ' style="list-style-type:circle;"'
                }
                Disc
                {
                    $inlineStyle = ' style="list-style-type:disc;"'
                }
                Square
                {
                    $inlineStyle = ' style="list-style-type:square;"'
                }
                Default
                {
                    ## Dash style is not supported in Html so default to the browser's rendering engine
                    $inlineStyle = ''
                }
            }

            $outHtmlListParams = @{
                BulletStyle = $BulletStyle
            }

            if ($List.HasStyle)
            {
                [ref] $null = $listBuilder.AppendFormat('{0}<ul class="{1}"{2}>', $leader, $List.Style, $inlineStyle)
            }
            else
            {
                [ref] $null = $listBuilder.AppendFormat('{0}<ul{1}>', $leader, $inlineStyle)
            }
            [ref] $null = $listBuilder.AppendLine()
        }

        $leader = ''.PadRight($Indent, ' ')
        foreach ($item in $List.Items)
        {
            if ($item.Type -eq 'PScribo.Item')
            {
                if (($item.HasStyle -eq $true) -and ($item.HasInlineStyle -eq $true))
                {
                    $inlineStyle = Get-HtmlListItemInlineStyle -Item $item
                    [ref] $null = $listBuilder.AppendFormat('{0}<li class="{1}" style="{2}">{3}</li>', $leader, $item.Style, $inlineStyle, $item.Text)
                }
                elseif ($item.HasStyle)
                {
                    [ref] $null = $listBuilder.AppendFormat('{0}  <li class="{1}"">{2}</li>', $leader, $item.Style, $item.Text)
                }
                elseif ($item.HasInlineStyle)
                {
                    $inlineStyle = Get-HtmlListItemInlineStyle -Item $item
                    [ref] $null = $listBuilder.AppendFormat('{0}  <li style="{1}">{2}</li>', $leader, $inlineStyle, $item.Text)
                }
                else
                {
                    [ref] $null = $listBuilder.AppendFormat('{0}  <li>{1}</li>', $leader, $item.Text)
                }
                [ref] $null = $listBuilder.AppendLine()
            }
            else
            {
                $nestedList = Out-HtmlList -List $item -Indent ($Indent +2) @outHtmlListParams
                [ref] $null = $listBuilder.AppendLine($nestedList)
            }
        }

        if ($List.IsNumbered)
        {
            [ref] $null = $listBuilder.AppendFormat('{0}</ol>', $leader)
        }
        else
        {
            [ref] $null = $listBuilder.AppendFormat('{0}</ul>', $leader)
        }
        [ref] $null = $listBuilder.AppendLine()

        return $listBuilder.ToString()
    }
}

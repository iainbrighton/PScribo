$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Get-HtmlTableStyle' {

        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { };

        Context 'By Named Parameter' {

            It 'creates default table style' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle

                $padding = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'padding:*'
                ($padding.Split(':').Trim())[1] | Should BeExactly '0.08rem 0.33rem 0rem 0.33rem'
                #$borderColor = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*'
                #($borderColor.Split(':').Trim())[1] | Should BeExactly '#000'
                #$borderWidth = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-width:*'
                #($borderWidth.Split(':').Trim())[1] | Should BeExactly '0em'
                $borderStyle = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*'
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'none'
                $borderCollapse = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-collapse:*'
                ($borderCollapse.Split(':').Trim())[1] | Should BeExactly 'collapse'
            }

            It 'creates custom table padding style of 5pt, 10pt, 5pt and 10pt' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -PaddingTop 5 -PaddingRight 10 -PaddingBottom 5 -PaddingLeft 10

                $padding = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'padding:*'
                ($padding.Split(':').Trim())[1] | Should BeExactly '0.42rem 0.83rem 0.42rem 0.83rem'
            }

            It 'creates custom table border color style when -BorderWidth is specified' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -BorderColor CcC -BorderWidth 1

                $borderColor = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*'
                ($borderColor.Split(':').Trim())[1] | Should BeExactly '#ccc'
                $borderStyle = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*'
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid'
            }

            It 'creates custom table border width style of 3pt' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -BorderWidth 3

                $borderWidth = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-width:*'
                ($borderWidth.Split(':').Trim())[1] | Should BeExactly '0.25rem'
                $borderStyle = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*'
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid'
            }

            It 'creates custom table border with no color style when no -BorderWidth specified' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -BorderColor '#aAaAaA'

                $borderStyle = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*'
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'none'
            }

            It 'creates custom table border color style' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -BorderColor '#aAaAaA' -BorderWidth 2

                $borderColor = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-color:*'
                ($borderColor.Split(':').Trim())[1] | Should BeExactly '#aaaaaa'
                $borderStyle = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'border-style:*'
                ($borderStyle.Split(':').Trim())[1] | Should BeExactly 'solid'
            }

            It 'centers table' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -Align Center

                $marginLeft = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-left:*'
                $marginRight = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-right:*'
                ($marginLeft.Split(':').Trim())[1] | Should BeExactly 'auto'
                ($marginRight.Split(':').Trim())[1] | Should BeExactly 'auto'
            }

            It 'aligns table to the right' {
                Style -Name Default -Font Helvetica -Default
                TableStyle -Name TestTableStyle -Align Right

                $marginLeft = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-left:*'
                $marginRight = ((Get-HtmlTableStyle -TableStyle $pscriboDocument.TableStyles['TestTableStyle']).Split(';').Trim()) -like 'margin-right:*'
                ($marginLeft.Split(':').Trim())[1] | Should BeExactly 'auto'
                ($marginRight.Split(':').Trim())[1] | Should BeExactly '0'
            }

        } #end context By Named Parameter
    }
}

$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$pluginRoot  = Split-Path -Path $here -Parent;
$testRoot  = Split-Path -Path $pluginRoot -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'Plugins\Html\Get-HtmlStyle' {

        ## Scaffold new document to initialise options/styles
        $pscriboDocument = Document -Name 'Test' -ScriptBlock { }

        Context 'By named parameter' {

            It 'creates single font default style' {
                Style -Name Test -Font Helvetica
                $fontFamily = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-family:*'
                ($fontFamily.Split(':').Trim())[1] | Should BeExactly "'Helvetica'"
                $fontSize = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*'
                ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92rem'
                $fontWeight = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-weight:*'
                ($fontWeight.Split(':').Trim())[1] | Should BeExactly 'normal'
                $fontStyle = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-style:*'
                $fontStyle | Should BeNullOrEmpty
                $textDecoration = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-decoration:*'
                $textDecoration | Should BeNullOrEmpty
                $textAlign = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*'
                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'left'
                $color = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*'
                ($color.Split(':').Trim())[1] | Should BeExactly '#000000'
                $backgroundColor = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*'
                $backgroundColor | Should BeNullOrEmpty
            }

            It 'uses invariant culture font size (#6)' {
                Style -Name Test -Font Helvetica
                $currentCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture.Name
                [System.Threading.Thread]::CurrentThread.CurrentCulture = 'da-DK'

                $fontSize = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*'

                [System.Threading.Thread]::CurrentThread.CurrentCulture = $currentCulture
                ($fontSize.Split(':').Trim())[1] | Should BeExactly '0.92rem'
            }

            It 'creates multiple font default style' {
                Style -Name Test -Font Helvetica,Arial,Sans-Serif

                $fontFamily = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-family:*'

                ($fontFamily.Split(':').Trim())[1] | Should BeExactly "'Helvetica','Arial','Sans-Serif'"
            }

            It 'creates single 12pt font' {
                Style -Name Test -Font Helvetica -Size 12

                $fontSize = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-size:*'

                ($fontSize.Split(':').Trim())[1] | Should BeExactly '1.00rem'
            }

            It 'creates bold font style' {
                Style -Name Test -Font Helvetica -Bold

                $fontWeight = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-weight:*'

                ($fontWeight.Split(':').Trim())[1] | Should BeExactly 'bold'
            }

            It 'creates center aligned font style' {
                Style -Name Test -Font Helvetica -Align Center

                $textAlign = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*'

                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'center'
            }

            It 'creates right aligned font style' {
                Style -Name Test -Font Helvetica -Align Right

                $textAlign = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*'

                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'right'
            }

            It 'creates justified font style' {
                Style -Name Test -Font Helvetica -Align Justify

                $textAlign = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-align:*'

                ($textAlign.Split(':').Trim())[1] | Should BeExactly 'justify'
            }

            It 'creates underline font style' {
                Style -Name Test -Font Helvetica -Underline

                $textDecoration = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'text-decoration:*'

                ($textDecoration.Split(':').Trim())[1] | Should BeExactly 'underline'
            }

            It 'creates italic font style' {
                Style -Name Test -Font Helvetica -Italic

                $fontStyle = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'font-style:*'

                ($fontStyle.Split(':').Trim())[1] | Should BeExactly 'italic'
            }

            It 'creates colored font style' {
                Style -Name Test -Font Helvetica -Color ABC

                $color = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*'

                ($color.Split(':').Trim())[1] | Should BeExactly '#abc'
            }

            It 'creates colored font style with #' {
                Style -Name Test -Font Helvetica -Color '#ABC'

                $color = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'color:*'

                ($color.Split(':').Trim())[1] | Should BeExactly '#abc'
            }

            It 'creates background colored font style' {
                Style -Name Test -Font Helvetica -BackgroundColor '#DEF'

                $backgroundColor = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*'

                ($backgroundColor.Split(':').Trim())[1] | Should BeExactly '#def'
            }

            It 'creates background colored font without #' {
                Style -Name Test -Font Helvetica -BackgroundColor 'DEF'

                $backgroundColor = ((Get-HtmlStyle -Style $pscriboDocument.Styles['Test']).Split(';').Trim()) -like 'background-color:*'

                ($backgroundColor.Split(':').Trim())[1] | Should BeExactly '#def'
            }

        } #end context By Named Parameter
    }
}
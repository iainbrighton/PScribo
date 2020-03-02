$here = Split-Path -Path $MyInvocation.MyCommand.Path -Parent;
$testRoot  = Split-Path -Path $here -Parent;
$moduleRoot = Split-Path -Path $testRoot -Parent;
Import-Module "$moduleRoot\PScribo.psm1" -Force;

InModuleScope 'PScribo' {

    Describe 'TableStyle' {

        $pscriboDocument = Document 'ScaffoldDocument' {}

        Context 'By named parameter' {

            It 'sets table style using named -Id, -HeaderStyle and -RowStyle parameters' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal

                $pscriboDocument.TableStyles['Test Table Style'].Name | Should BeExactly 'Test Table Style'
                $pscriboDocument.TableStyles['Test Table Style'].Id | Should BeExactly 'TestTableStyle'
                $pscriboDocument.TableStyles['Test Table Style'].HeaderStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].RowStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].AlternateRowStyle | Should BeExactly 'Normal'
            }

            It 'sets table style using named -Id, -HeaderStyle, -RowStyle and -AlternateRowStyle parameters' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -AlternateRowStyle Normal

                $pscriboDocument.TableStyles['Test Table Style'].Name | Should BeExactly 'Test Table Style'
                $pscriboDocument.TableStyles['Test Table Style'].Id | Should BeExactly 'TestTableStyle'
                $pscriboDocument.TableStyles['Test Table Style'].HeaderStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].RowStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].AlternateRowStyle | Should BeExactly 'Normal'
            }

            It 'defaults table style to left alignment' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal

                $pscriboDocument.TableStyles['Test Table Style'].Align | Should BeExactly 'Left'
            }

            It 'sets table style alignment to center by named -Align parameter' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -Align Center

                $pscriboDocument.TableStyles['Test Table Style'].Align | Should BeExactly 'Center'
            }

            It 'sets table style alignment to right by named -Align parameter' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -Align Right

                $pscriboDocument.TableStyles['Test Table Style'].Align | Should BeExactly 'Right'
            }

            It 'defaults table style padding to 1.0pt, 4.0pt, 1.0pt and 4.0pt' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal

                $pscriboDocument.TableStyles['Test Table Style'].PaddingTop | Should Be 0.35
                $pscriboDocument.TableStyles['Test Table Style'].PaddingRight | Should Be 1.41
                $pscriboDocument.TableStyles['Test Table Style'].PaddingBottom | Should Be 0.0
                $pscriboDocument.TableStyles['Test Table Style'].PaddingLeft | Should Be 1.41
            }

            It 'sets table style padding to 2.0pt, 5.0pt, 1.0pt and 5.0pt' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -PaddingTop 2.0 -PaddingRight 5.0 -PaddingBottom 1.0 -PaddingLeft 5.0

                $pscriboDocument.TableStyles['Test Table Style'].PaddingTop | Should Be 0.71
                $pscriboDocument.TableStyles['Test Table Style'].PaddingLeft | Should Be 1.76
                $pscriboDocument.TableStyles['Test Table Style'].PaddingRight | Should Be 1.76
                $pscriboDocument.TableStyles['Test Table Style'].PaddingBottom | Should Be 0.35
            }

            It 'defaults table style border to none' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal

                $pscriboDocument.TableStyles['Test Table Style'].BorderStyle | Should BeExactly 'None'
                $pscriboDocument.TableStyles['Test Table Style'].BorderWidth | Should Be 0
            }

            It 'sets table style border to 2pt by named -BorderWidth parameter' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -BorderWidth 2

                $pscriboDocument.TableStyles['Test Table Style'].BorderStyle | Should BeExactly 'Solid'
                $pscriboDocument.TableStyles['Test Table Style'].BorderWidth | Should Be 0.71
            }

            It 'sets table style border to 0px by named -BorderWidth parameter' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -BorderWidth 0

                $pscriboDocument.TableStyles['Test Table Style'].BorderStyle | Should BeExactly 'None'
                $pscriboDocument.TableStyles['Test Table Style'].BorderWidth | Should Be 0
            }

            It 'sets Normal table style by named -Normal parameter' {
                TableStyle -Id 'Test Table Style' -HeaderStyle Normal -RowStyle Normal -Default

                $pscriboDocument.DefaultTableStyle | Should Be 'TestTableStyle'
            }

            It 'throws with invalid named -HeaderStyle parameter' {
                { TableStyle 'Test Table Style' -HeaderStyle InvalidStyle } | Should Throw
            }

            It 'throws with invalid named -RowStyle parameter' {
                { TableStyle 'Test Table Style' -RowStyle InvalidStyle } | Should Throw
            }

            It 'throws with invalid named -AlternateRowStyle parameter' {
                { TableStyle 'Test Table Style' -AlternateRowStyle InvalidStyle } | Should Throw
            }

            It 'throws with invalid named -BorderColor parameter' {
                { TableStyle 'Test Table Style' -BorderColor xyz } | Should Throw
            }
        }

        Context 'By positional parameter' {

            It 'sets table style using positional -Id, -HeaderStyle and -RowStyle parameters' {
                TableStyle 'Test Table Style' Normal Normal

                $pscriboDocument.TableStyles['Test Table Style'].Name | Should BeExactly 'Test Table Style'
                $pscriboDocument.TableStyles['Test Table Style'].Id | Should BeExactly 'TestTableStyle'
                $pscriboDocument.TableStyles['Test Table Style'].HeaderStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].RowStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].AlternateRowStyle | Should BeExactly 'Normal'
            }

            It 'sets table style using positional -Id, -HeaderStyle, -RowStyle and -AlternateRowStyle parameters' {
                TableStyle 'Test Table Style' Normal Normal Normal

                $pscriboDocument.TableStyles['Test Table Style'].Name | Should BeExactly 'Test Table Style'
                $pscriboDocument.TableStyles['Test Table Style'].Id | Should BeExactly 'TestTableStyle'
                $pscriboDocument.TableStyles['Test Table Style'].HeaderStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].RowStyle | Should BeExactly 'Normal'
                $pscriboDocument.TableStyles['Test Table Style'].AlternateRowStyle | Should BeExactly 'Normal'
            }

            It 'throws with invalid positional -HeaderStyle parameter' {
                { TableStyle 'Test Table Style' InvalidStyle Normal } | Should Throw
            }

            It 'throws with invalid positional -RowStyle parameter' {
                { TableStyle 'Test Table Style' Normal InvalidStyle } | Should Throw
            }

            It 'throws with invalid positional -AlternateRowStyle parameter' {
                { TableStyle 'Test Table Style' Normal Normal InvalidStyle } | Should Throw
            }
        }
    }
}

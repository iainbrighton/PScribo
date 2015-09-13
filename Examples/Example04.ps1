Import-Module PScribo -Force;

$example4 = Document -Name 'PScribo Example 4' {
    Paragraph 'The default built-in style that is applied to paragraphs is named "Normal".'
    Paragraph 'You can apply an alternative color to a particular paragraph by specifying a html color code to the -Color parameter.' -Color f00;
    Paragraph 'You can also supply Word constant colors to a paragraph!' -Color SteelBlue;
    Paragraph 'You can apply bold styling to a particular paragraph by specifying the -Bold switch parameter.' -Bold
    Paragraph 'You can apply italic styling to a particular paragraph by specifying the -Italic switch parameter.' -Italic
    Paragraph 'You can apply underline styling to a particular paragraph by specifying the -Underline switch parameter.' -Underline
    Paragraph 'You can alter the default font size of a particular paragraph by specifying the font point size with the -Size parameter' -Size 16
    Paragraph 'You can alter the default font typeface of a particular paragraph by specifying -Font parameter. This parameter takes an array of font names. Typically, only the first font name is used in Word output, but all names are used in Html output.' -Font Arial
    Paragraph 'If you wish to indent a paragraph, specify the -Tabs parameter. This is an integer number that indents a paragraph at 12.7mm/0.5inch intervals.' -Tabs 1
    Paragraph 'Of course you are free to combine multiple styling parameters to a single paragraph :).' -Color DarkOrchid -Size 14 -Bold
}
$example4 | Export-Document -Format Html -Path ~\Desktop

function Out-MarkdownImage
{
<#
    .SYNOPSIS
        Output formatted image text.
#>
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Image
    )
    begin
    {
        ## Fix Set-StrictMode
        if (-not (Test-Path -Path Variable:\Options))
        {
            $options = New-PScriboMarkdownOption
        }
    }
    process
    {

        $imageBuilder = New-Object -TypeName System.Text.StringBuilder
        if ($options.EmbedImage)
        {
            [ref] $null = $imageBuilder.AppendFormat('![{0}]({1})', $Image.Text, $Image.Uri).AppendLine().AppendLine()
        }
        else
        {
            $imageUri = $Image.Uri -as [System.Uri]
            if ($imageUri.IsFile)
            {
                [ref] $null = $imageBuilder.AppendFormat('![{0}]({1})', $Image.Text, $imageUri.LocalPath).AppendLine().AppendLine()
            }
            else
            {
                [ref] $null = $imageBuilder.AppendFormat('![{0}]({1})', $Image.Text, $imageUri.AbsoluteUri).AppendLine().AppendLine()
            }
        }
        return $imageBuilder.ToString()
        <#
        ![][image_ref_a32ff4ads]

More text here...
...

[image_ref_a32ff4ads]: data:image/png;base64,iVBORw0KGgoAAAANSUhEke02C1MyA29UWKgPA...RS12D==
        #>
    }
}

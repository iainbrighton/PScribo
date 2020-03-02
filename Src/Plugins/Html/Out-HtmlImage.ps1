function Out-HtmlImage
{
<#
    .SYNOPSIS
        Output embedded Html image.
#>
    [CmdletBinding()]
    param
    (
        ## PScribo Image object
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [System.Management.Automation.PSObject] $Image
    )
    process
    {
        [System.Text.StringBuilder] $imageBuilder = New-Object -TypeName 'System.Text.StringBuilder'
        [ref] $null = $imageBuilder.AppendFormat('<div align="{0}">', $Image.Align).AppendLine()
        $imageBase64 = [System.Convert]::ToBase64String($Image.Bytes)
        [ref] $null = $imageBuilder.AppendFormat('<img src="data:{0};base64, {1}" alt="{2}" height="{3}" width="{4}" />', $Image.MimeType, $imageBase64, $Image.Text, $Image.Height, $Image.Width).AppendLine()
        [ref] $null = $imageBuilder.AppendLine('</div>')
        return $imageBuilder.ToString()
    }
}

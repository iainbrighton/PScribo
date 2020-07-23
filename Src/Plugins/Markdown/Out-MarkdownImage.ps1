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
        if (($null -ne $Options) -and ($Options['EmbedImage'] -eq $true))
        {
            [ref] $null = $imageBuilder.AppendFormat('![{0}][ref_{1}]', $Image.Text, $Image.Name.ToLower()).AppendLine().AppendLine()
        }
        else
        {
            $imageUri = $Image.Uri -as [System.Uri]
            [ref] $null = $imageBuilder.AppendFormat('![{0}]({1})', $Image.Text, $imageUri.AbsoluteUri).AppendLine().AppendLine()
        }
        $script:currentPScriboObject = 'PScribo.Image'
        return $imageBuilder.ToString()
    }
}

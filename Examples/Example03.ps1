[CmdletBinding()]
param (
    [System.String[]] $Format = 'Html',
    [System.String] $Path = '~\Desktop',
    [System.Management.Automation.SwitchParameter] $PassThru
)

Import-Module PScribo -Force -Verbose:$false

Document -Name 'PScribo Example 3' {

    <#
        By default, a PScribo document will default to an A4 page size. The default top and bottom
        pages margins are configured at 25.4mm (or 1 inch). The default left and right page margins
        are set at 19.05mm (or 3/4 inch).

        You can override the defaults with the 'DocumentOption' cmdlet/keyword. The following sets the
        page size to US Letter.
    #>
    DocumentOption -PageSize Letter

    <#
        The page margins are specified in points (pt). The follow cmdlet sets all page margins to
        12.7mm (or 0.5 inch).
    #>
    DocumentOption -Margin 36

    <#
        The top and bottom page margins can be set independently from the left and right margin. The
        following configures the top/bottom page margin to 19.05mm (3/4 inch) and the left/right page
        margin to 12.7mm (0.5 inch).
    #>
    DocumentOption -MarginTopAndBottom 54 -MarginLeftAndRight 36

    <#
        Multiple options can be specified in a single call to 'DocumentOption', e.g. setting the
        page size and margins like so:
    #>
    DocumentOption -PageSize Letter -MarginTopAndBottom 54 -MarginLeftAndRight 36

} | Export-Document -Path $Path -Format $Format -PassThru:$PassThru

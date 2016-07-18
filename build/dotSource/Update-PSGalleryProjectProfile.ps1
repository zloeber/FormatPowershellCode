#Requires -version 5
function Script:Update-PSGalleryProjectProfile {
    <#
        .SYNOPSIS
            Update a powershell Gallery module upload profile
        .DESCRIPTION
            Update a powershell Gallery module upload profile
        .PARAMETER Name
            Module short name.
        .PARAMETER Path
            Path of module project files to upload.
        .PARAMETER ProjectUri
            Module project website.
        .PARAMETER Tags
            Tags used to search for the module (separated by spaces)
        .PARAMETER RequiredVersion
            Module version
        .PARAMETER Repository
            Destination gallery (default is PSGallery)
        .PARAMETER ReleaseNotes
            Release notes.
        .PARAMETER LicenseUri
            License website.
        .PARAMETER IconUri
            Icon web path.
        .PARAMETER APIKey
            API key for the powershellgallery.com site. 
        .PARAMETER OutputFile
            Input module configuration file (default is .psgallery)

        .EXAMPLE
        .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Version History
        1.0.0 - Initial release
        #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Module short name.')]
        [string]$Name,
        [parameter(Position=1, HelpMessage='Path of module project files to upload.')]
        [string]$Path,
        [parameter(Position=2, HelpMessage='Module project website.')]
        [string]$ProjectUri,
        [parameter(Position=3, HelpMessage='Tags used to search for the module (separated by spaces)')]
        [string]$Tags,
        [parameter(Position=4, HelpMessage='Required powershell version (default is 2)')]
        [string]$RequiredVersion,
        [parameter(Position=5, HelpMessage='Destination gallery (default is PSGallery)')]
        [string]$Repository,
        [parameter(Position=6, HelpMessage='Release notes.')]
        [string]$ReleaseNotes,
        [parameter(Position=7, HelpMessage='License website.')]
        [string]$LicenseUri,
        [parameter(Position=9, HelpMessage='Icon web path.')]
        [string]$IconUri,
        [parameter(Position=10, HelpMessage='API key for the powershellgallery.com site.')]
        [string]$NuGetApiKey,
        [parameter(Position=11, HelpMessage='Input module configuration file (default is .psgallery)')]
        [string]$InputFile = '.psgallery'
    )

    if (Test-Path $InputFile) {
        $PublishParams  = Import-Clixml $InputFile
        $MyParams = $PSCmdlet.MyInvocation.BoundParameters
        $MyParams.Keys | Where {$_ -ne 'InputFile'} | ForEach {
            Write-Verbose "Updating $($_)"
            if ($PublishParams.$_ -ne $null) {
                $PublishParams.$_ = $MyParams[$_]
            }
        }
        $PublishParams | Export-Clixml -Path $InputFile -Force
    }
    else {
        Write-Warning "InputFile was not found: $($InputFile)"
    }
}
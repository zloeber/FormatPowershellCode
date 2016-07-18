#Requires -version 5
function Script:New-PSGalleryProjectProfile {
    <#
        .SYNOPSIS
            Create a powershell Gallery module upload profile
        .DESCRIPTION
            Create a powershell Gallery module upload profile
        .PARAMETER Name
            Module short name.
        .PARAMETER Path
            Path of module project files to upload.
        .PARAMETER ProjectUri
            Module project website.
        .PARAMETER Tags
            Tags used to search for the module (separated by spaces)
        .PARAMETER RequiredVersion
            Required powershell version (default is 2)
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
            OutputFile (default is .psgallery)

        .EXAMPLE
        .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Version History
        1.0.0 - Initial release
        #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, Mandatory=$true, HelpMessage='Module short name.')]
        [string]$Name,
        [parameter(Position=1, Mandatory=$true, HelpMessage='Path of module project files to upload.')]
        [string]$Path,
        [parameter(Position=2, HelpMessage='Module project website.')]
        [string]$ProjectUri = '',
        [parameter(Position=3, HelpMessage='Tags used to search for the module (separated by spaces)')]
        [string]$Tags = '',
        [parameter(Position=4, HelpMessage='Required powershell version (default is 2)')]
        [string]$RequiredVersion = 2,
        [parameter(Position=5, HelpMessage='Destination gallery (default is PSGallery)')]
        [string]$Repository = 'PSGallery',
        [parameter(Position=6, HelpMessage='Release notes.')]
        [string]$ReleaseNotes = '',
        [parameter(Position=7, HelpMessage=' License website.')]
        [string]$LicenseUri = '',
        [parameter(Position=9, HelpMessage='Icon web path.')]
        [string]$IconUri = '',
        [parameter(Position=10, HelpMessage='API key for the powershellgallery.com site.')]
        [string]$APIKey = '',
        [parameter(Position=11, HelpMessage='OutputFile (default is .psgallery)')]
        [string]$OutputFile = '.psgallery'
    )

    $PublishParams = @{
        Name = $Name
        Path = $Path
        APIKey = $APIKey
        ProjectUri = $ProjectUri
        Tags = $Tags
        RequiredVersion = $RequiredVersion
        Repository = $Repository
        ReleaseNotes = $ReleaseNotes
        LicenseUri = $LicenseUri
        IconUri = $IconUri
    }

    if (Test-Path $OutputFile) {
        $PublishParams | Export-Clixml -Path $OutputFile -confirm
    }
    else {
        $PublishParams | Export-Clixml -Path $OutputFile
    }
}
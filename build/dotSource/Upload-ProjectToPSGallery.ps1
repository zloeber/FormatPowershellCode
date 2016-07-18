#Requires -version 5
function Script:Upload-ProjectToPSGallery {
    <#
        .SYNOPSIS
            Upload module project to Powershell Gallery
        .DESCRIPTION
            Upload module project to Powershell Gallery
        .PARAMETER ModulePath
            Path to module to upload.
        .PARAMETER APIKey
            API key for the powershellgallery.com site. 
        .PARAMETER Tags
            Tags for your module
        .PARAMETER ProjectURI
            Project site (like github).
        .EXAMPLE
            .\Upload-ProjectToPSGallery.ps1
        .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Requires: Powershell 5.0

        Version History
        1.0.0 - Initial release
        #>
    [CmdletBinding(DefaultParameterSetName='PSGalleryProfile')]
    param(
        [parameter(Mandatory=$true, 
                        HelpMessage='Module short name.',
                        ParameterSetName='ManualInput')]
        [string]$Name,
        [parameter(Mandatory=$true, 
                        HelpMessage='Path of module project files to upload.',
                        ParameterSetName='ManualInput')]
        [string]$Path,
        [parameter(HelpMessage='Module project website.',
                        ParameterSetName='ManualInput')]
        [string]$ProjectUri,
        [parameter(HelpMessage='Tags used to search for the module (separated by spaces)',
                        ParameterSetName='ManualInput')]
        [string]$Tags,
        [parameter(HelpMessage='Required powershell version (default is 2)',
                        ParameterSetName='ManualInput')]
        [string]$RequiredVersion,
        [parameter(HelpMessage='Destination gallery (default is PSGallery)',
                        ParameterSetName='ManualInput')]
        [string]$Repository = 'PSGallery',
        [parameter(HelpMessage='Release notes.',
                        ParameterSetName='ManualInput')]
        [string]$ReleaseNotes,
        [parameter(HelpMessage=' License website.',
                        ParameterSetName='ManualInput')]
        [string]$LicenseUri,
        [parameter(HelpMessage='Icon web path.',
                        ParameterSetName='ManualInput')]
        [string]$IconUri,
        [parameter(Mandatory = $true,
                        HelpMessage='API key for the powershellgallery.com site.',
                        ParameterSetName='ManualInput')]
        [parameter(HelpMessage='API key for the powershellgallery.com site.',
                        ParameterSetName='PSGalleryProfile')]
        [string]$NuGetApiKey,

        [parameter(HelpMessage='Path to CliXML file containing your psgallery project information.',
                        ParameterSetName='PSGalleryProfile')]
        [string]$PSGalleryProfilePath = '.psgallery'
    )

    Write-Verbose "Using parameterset $($PSCmdlet.ParameterSetName)"
    if ($PSCmdlet.ParameterSetName -eq 'PSGalleryProfile') {
        if (Test-Path $PSGalleryProfilePath) {
            Write-Verbose "Loading PSGallery profile information from $PSGalleryProfilePath"
            $PublishParams = Import-Clixml $PSGalleryProfilePath
        }
        else {
            Write-Error "$PSGalleryProfilePath not found"
            return
        }
    }
    else {
        $MyParams = $PSCmdlet.MyInvocation.BoundParameters
        $MyParams.Keys | ForEach {
            Write-Verbose "Adding manually defined parameter $($_)"
            $PublishParams.$_ = $MyParams[$_]
        }
    }

    # if no API key is defined then look for psgalleryapi.txt in the local profile directory and try to use it instead.
    if ([string]::IsNullOrEmpty($PublishParams.NuGetApiKey)) {
        $psgalleryapipath = "$(Split-Path $Profile)\psgalleryapi.txt"
        Write-Verbose "No PSGallery API key specified. Attempting to load one from the following location: $($psgalleryapipath)"
        if (-not (test-path $psgalleryapipath)) {
            Write-Error "$psgalleryapipath wasn't found and there was no defined API key, please rerun script with a defined APIKey parameter."
            return
        }
        else {
            $PublishParams.NuGetApiKey = get-content -raw $psgalleryapipath
        }
    }

    $NewTags = $PublishParams.Tags -replace ',','","'
    $PublishParams.Tags = $PublishParams.Tags -split ','

    # If we made it this far then try to publish the module wth our loaded parameters
    Publish-Module @PublishParams
}
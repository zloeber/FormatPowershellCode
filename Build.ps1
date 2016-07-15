#Requires -Version 5

<#
.Synopsis
	Build script using Invoke-Build (https://github.com/nightroman/Invoke-Build)
.Description
	The script automates the build process for the FormatPowerShellCode module. 
    
    The overarching build steps are:
    * Clean up the build temp directory
    * In the temp directory: 
        - Create project folder structure
        - Copy over project files
        - Format PowerShell files with the FormatPowerShellCode module
        - Create module markdown files (PlatyPS)
        - Convert markdown files to HTML documents for online help (PlatyPS)
        - Create and automatically fill in the blanks for the online help landing page (PlatyPS)
        - Create the online help cab download (PlatyPS)
        - Update module manifest with exported functions and new version.
        - Combine the existing public and private PowerShell script files into one psm1

    Some additional or planned build steps are:
    * Automatically kick of Pester testing
	* Push the release with a tag to GitHub
    * Push the release to PSGallery

    This build script has several dependencies which are delt with in the 'Configure' task
        Invoke-Build - Runs build tasks
        PlatyPS - Create html help files
        FormatPowerShellCode - Pretty up the code (This build script is for this same module so not really a dependency but still...)
        .psgallery - File containing information for pushing module to psgallery
        version.txt - Used to determine version of build to create (all versions will overwrite the 'current' release folder)

    To run the build:
        .\build.ps1

    Notes:
    - The public functions are based on the name of the files within the /src/public folder
    - This build is based on your existing module being loaded as it is and will infer information from it to build the final release
    - The release number is driven by the version.txt file in the root of your module project directory. You can update the exising module
      manifest in this directory with this version with:
        Invoke-Build UpdateVersion
    - I use powershellget to ease installation of required modules. This will have to be rewritten in several spots to attain any kind of 
      backward compatibility.
    - There is no real accounting for exported variables, aliases, or other public content in this script. The manifest will copy over manually defined
      items like this but you 

#>
if ((get-module InvokeBuild -ListAvailable) -eq $null) {
    Write-Host -NoNewLine "      Installing InvokeBuild module"
    $null = Install-Module InvokeBuild
    Write-Host -ForegroundColor Green '...Installed!'
}
if (get-module InvokeBuild -ListAvailable) {
    Write-Host -NoNewLine "      Importing InvokeBuild module"
    Import-Module InvokeBuild -Force
    Write-Host -ForegroundColor Green '...Loaded!'
}
else {
    throw 'How did you even get here?'
}

# Kick off the standard build
try {
    Invoke-Build
}
catch {
    # If it fails then show the error and try to clean up the environment
    Write-Host -ForegroundColor Red 'Build Failed with the following error:'
    Write-Output $_
}
finally {
    Write-Host ''
    Write-Host 'Attempting to clean up the session (loaded modules and such)...'
    Invoke-Build BuildSessionCleanup
    Remove-Module InvokeBuild
}
<#
.Synopsis
	Build script using Invoke-Build (https://github.com/nightroman/Invoke-Build)
.Description
	The script automates the build process for the FormatPowerShellCode module. The overarching steps are:
    * Clean the build temp directory
    * In the temp directory: 
        - Create project folder structure
        - Copy over project files
        - Format PowerShell files with the FormatPowerShellCode module
        - Create module markdown files (PlatyPS)
        - Convert markdown files to HTML documents for online help (PlatyPS)
        - Update module manifest with exported functions and new version.

	* Push the release with a tag to GitHub
    * Push the release to PSGallery

    This build script has several dependencies:
        Invoke-Build - Runs build tasks
        PlatyPS - Create html help files
        FormatPowerShellCode - Pretty up the code
        .psgallery - File containing information for pushing module to psgallery
        version.txt - Used to determine version of build to create (all versions will overwrite the 'current' release folder)

    To run the build:
    Ensure PlatyPS module is available to import
    Either get a copy of invoke-build.ps1 locally or install the module (install-module invokebuild) then just run,

        invoke-build

#>

$ModuleToBuild = 'FormatPowershellCode' 
$ModuleFullPath = (Get-Item "$($ModuleToBuild).psm1").FullName
#$ModuleManifestFullPath = (Get-Item "$($ModuleToBuild).psd1").FullName
$ScriptRoot = Split-Path $ModuleFullPath
$TempPath = "$($ScriptRoot)\temp"
$StageReleasePath = Join-Path $TempPath 'release'
#$BuildPath = "$($ScriptRoot)\build"
$VersionFile = "$($ScriptRoot)\version.txt"
$ReleasePath = "$($ScriptRoot)\release"
$CurrentReleasePath = "$($ScriptRoot)\release\current"

# Some extra functions
function Out-Zip {
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [string] $Directory,
        [Parameter(Position=1, Mandatory=$true)]
        [string] $FileName,
        [Parameter(Position=2)]
        [switch] $overwrite
    )
    Add-Type -Assembly System.IO.Compression.FileSystem
    $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
    if (-not $FileName.EndsWith('.zip')) {$FileName += '.zip'} 
    if ($overwrite) {
        if (Test-Path $FileName) {
            Remove-Item $FileName
        }
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($Directory, $FileName, $compressionLevel, $false)
}

<# Synopsis: Validate script requirements are met and load build tools.
task Configure {
    # Dot source any build script functions we need to use
    Get-ChildItem $BuildPath/dotSource -Recurse -Filter "*.ps1" -File | Foreach { 
        Write-Verbose "Dot sourcing script file: $($_.Name)"
        . $_.FullName
    }
}
#>

# Synopsis: Set $script:Version.
task Version {
    $script:Version = [version](Get-Content $VersionFile)
    #$ModVer = .{switch -Regex -File $ModuleManifestFullPath {"ModuleVersion\s+=\s+'(\d+\.\d+\.\d+)'" {return $Matches[1]}}}
    $ModVer = (Get-Module -ListAvailable -Name .\FormatPowerShellCode.psd1).Version.ToString()
    assert ($ModVer -eq (($Version).ToString())) "The module manifest version ($($ModVer)) and release version ($($Version)) are mismatched."
}

# Synopsis: Remove/regenerate temp staging directory
task Clean {
	Remove-Item $TempPath -Force -Recurse -ErrorAction 0
    New-Item $TempPath -ItemType:Directory
}

# Synopsis: Create base content tree in staging area
task PrepareStage {
    New-Item "$($TempPath)\src" -ItemType:Directory
    New-Item $StageReleasePath -ItemType:Directory

    Copy-Item -Path "$($ScriptRoot)\*.psm1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\*.psd1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\src\public" -Recurse -Destination "$($TempPath)\src"
    Copy-Item -Path "$($ScriptRoot)\src\private" -Recurse -Destination "$($TempPath)\src"
    Copy-Item -Path "$($ScriptRoot)\en-US" -Recurse -Destination $TempPath

    # Load the manifest data
    $Script:Manifest = Import-PowerShellDataFile -Path $TempPath\$ModuleToBuild.psd1
    
    # Collect a list of our public methods for later module manifest updates
    $Script:FunctionsToExport = (Get-ChildItem -Path $TempPath\src\public).BaseName
}

# Synopsis: Assemble the module
task PrepareModuleRelease {
    if(-not $Script:Manifest.RootModule) {
        $Script:Manifest.RootModule = $Manifest.ModuleToProcess
        if(-not $Script:Manifest.RootModule) {
            $Script:Manifest.RootModule = "$ModuleToBuild.psm1"
        }
    }
    $ReleaseModule = Join-Path $StageReleasePath $Manifest.RootModule
    
    Set-Content $ReleaseModule ((
        (Get-Content (Join-Path $TempPath 'src\private\*.ps1') -Raw) + 
        (Get-Content (Join-Path $TempPath 'src\public\*.ps1') -Raw)) -join "`r`n`r`n`r`n") -Encoding UTF8

    $Script:Manifest.ModuleVersion = $Script:Version
    $Script:Manifest.FunctionsToExport = $Script:FunctionsToExport

    # There is something wrong with -privatedata when passed into new-modulemanifest. this may not really work but it at least leaves a usable psd1
    $PrivateData = $null
    if ($Script:Manifest.Keys -contains 'PrivateData') {
        $PrivateData = New-Object -TypeName psobject -Property $Script:Manifest.PrivateData.PSdata
        $Script:Manifest.Remove('PrivateData')
    }
    $MyManifest = $Script:Manifest
    New-ModuleManifest @MyManifest -Path $StageReleasePath\$ModuleToBuild.psd1 -PrivateData $PrivateData

    # Copy over documentation last
    Copy-Item -Path "$($TempPath)\en-US" -Recurse -Destination $StageReleasePath
}

# Synopsis: Warn about not empty git status if .git exists.
task GitStatus -If (Test-Path .git) {
	$status = exec { git status -s }
	if ($status) {
		Write-Warning "Git status: $($status -join ', ')"
	}
}

# Synopsis: Run code formatter against our working build (dogfood)
task FormatCode {
    	Import-Module $ModuleFullPath
        Get-ChildItem -Path $TempPath -Include "*.ps1","*.psm1" -Recurse -File | Where {$_.FullName -notlike "$($StageReleasePath)*"} | ForEach {
            $FormattedOutFile = $_.FullName
            Write-Output "Formatting File: $($FormattedOutFile)"
            $FormattedCode = get-content $_ -raw |
                Format-ScriptRemoveStatementSeparators |
               # Format-ScriptExpandFunctionBlocks |
                Format-ScriptExpandNamedBlocks |
                Format-ScriptExpandParameterBlocks |
                Format-ScriptExpandStatementBlocks |
                Format-ScriptPadOperators |
                Format-ScriptPadExpressions |
                Format-ScriptFormatTypeNames |
                Format-ScriptReduceLineLength |
                Format-ScriptRemoveSuperfluousSpaces |
                Format-ScriptFormatCodeIndentation

                $FormattedCode | Out-File -FilePath $FormattedOutFile -force
        }
        Remove-Module $ModuleToBuild
}

# Synopsis: Run PSScriptAnalyzer against the module
task AnalyzeScripts {

}

# Synopsis: Update the version information in the gallery data upload file
task UpdatePSGalleryData {

}

# Synopsis: Push the project to PSScriptGallery
task PublishPSGallery {

}

<# Synopsis: Update module manifest with exported functions and version.
task UpdateModuleManifest -if ($Script:Manifest -ne $null) {
    $Script:Manifest.ModuleVersion = $Script:Version
    $Script:Manifest.FunctionsToExport = $Script:FunctionsToExport

    # There is something wrong with -privatedata when passed into new-modulemanifest. this may not really work but it at least leaves a usable psd1
    $PrivateData = $null
    if ($Script:Manifest.Keys -contains 'PrivateData') {
        $PrivateData = New-Object -TypeName psobject -Property $Script:Manifest.PrivateData.PSdata
        $Script:Manifest.Remove('PrivateData')
    }
    $MyManifest = $Script:Manifest
    New-ModuleManifest @MyManifest -Path $StageReleasePath\$ModuleToBuild.psd1 -PrivateData $PrivateData
} #>

# Synopsis: Create new release version directory from our temporary build directory and copy our results
task PublishModuleVersionRelease -If ( -not (Test-Path "$($ReleasePath)\$($Script:Version)")) {
    New-Item "$($ReleasePath)\$($Script:Version)" -ItemType:Directory
    Copy-Item $TempPath -Destination "$($ReleasePath)\$($Script:Version)" -Recurse
}

# Synopsis: Build the html help files and external module help files with PlatyPS
task CreateHelp {
	Import-Module PlatyPS
    
    # We need to import the module to give the helps functions something to work against
    Import-Module (Join-Path $StageReleasePath $Manifest.RootModule)

    # Create new html help files
    New-MarkdownHelp -module $ModuleToBuild -OutputFolder "$($StageReleasePath)\docs\" -Force
    
    # Create external help files
    New-ExternalHelp "$($StageReleasePath)\docs" -OutputPath "$($StageReleasePath)\en-US\"

    # Clean up loaded modules
    Remove-Module $ModuleToBuild
    Remove-Module PlatyPS
}

# Synopsis: Create a new version release directory for our release and copy our contents to it
task PushVersionRelease {
    $ThisReleasePath = Join-Path $ReleasePath $Script:Version
    Remove-Item $ThisReleasePath -Force -Recurse -ErrorAction 0
    New-Item $ThisReleasePath -ItemType:Directory
    Copy-Item -Path "$($StageReleasePath)\*" -Destination $ThisReleasePath -Recurse
    Out-Zip $StageReleasePath $ReleasePath\$ModuleToBuild'-'$Version'.zip' -overwrite
}

# Synopsis: Create the current release directory and copy this build to it.
task PushCurrentRelease {
    Remove-Item $CurrentReleasePath -Force -Recurse -ErrorAction 0
    New-Item $CurrentReleasePath -ItemType:Directory
    Copy-Item -Path "$($StageReleasePath)\*" -Destination $CurrentReleasePath -Recurse
    Out-Zip $StageReleasePath $ReleasePath\$ModuleToBuild'-current.zip' -overwrite
}
<# Synopsis: Push with a version tag.
task PushRelease Version, {
	$changes = exec { git status --short }
	assert (!$changes) "Please, commit changes."

	exec { git push }
	exec { git tag -a "v$Version" -m "v$Version" }
	exec { git push origin "v$Version" }
}

# Synopsis: Remove any script signatures which may be applied to the individual files
task RemoveScriptSignatures {
	 Get-ChildItem -Path $TempPath -File -Recurse | Remove-Signature
}
#>


# Synopsis: The default build
task . Version, Clean, PrepareStage, FormatCode, PrepareModuleRelease, CreateHelp, PushVersionRelease, PushCurrentRelease, GitStatus

# Synopsis: Test the code formatting module only
task TestModule Version, Clean, PrepareStage, FormatCode
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
$BuildPath = "$($ScriptRoot)\build"
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

# Synopsis: Validate script requirements are met and load build tools.
task Configure {
    # Dot source any build script functions we need to use
    Get-ChildItem $BuildPath/dotSource -Recurse -Filter "*.ps1" -File | Foreach { 
        Write-Verbose "Dot sourcing script file: $($_.Name)"
        . $_.FullName
    }

}
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
task Stage {
    New-Item "$($TempPath)\src" -ItemType:Directory
    Copy-Item -Path "$($ScriptRoot)\*.psm1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\*.psd1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\version.txt" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\Install.ps1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\*.md" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\src\public" -Recurse -Destination "$($TempPath)\src"
    Copy-Item -Path "$($ScriptRoot)\src\private" -Recurse -Destination "$($TempPath)\src"
    Copy-Item -Path "$($ScriptRoot)\en-US" -Recurse -Destination $TempPath

    # Collect a list of our public methods for later module manifest updates
    $Script:FunctionsToExport = (Get-ChildItem -Path $TempPath\src\public).BaseName

    # Load the manifest data
    $Script:Manifest = Import-PowerShellDataFile -Path $TempPath\$ModuleToBuild.psd1
}

# Synopsis: Combine module into one script for staging
task StageCombine {
    Copy-Item -Path "$($ScriptRoot)\*.psd1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\version.txt" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\Install.ps1" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\*.md" -Destination $TempPath
    Copy-Item -Path "$($ScriptRoot)\en-US" -Recurse -Destination $TempPath

    $Manifest = Import-PowerShellDataFile -Path $TempPath\$ModuleToBuild.psd1

    # if the Source folder has "Public" and optionally "Private" in it, then the psm1 must be assembled:
    if(Test-Path (Join-Path $ScriptRoot 'src\Public') -Type Container) {
        $Script:FunctionsToExport = (Get-ChildItem -Path (Join-Path $ScriptRoot 'src\Public')).BaseName
        #Write-Verbose "       Collating Module Source"

        if(!$Manifest.RootModule) {
            $Manifest.RootModule = $Manifest.ModuleToProcess
            if(!$Manifest.RootModule) {
                $Manifest.RootModule = "$ModuleToBuild.psm1"
            }
        }

        $ReleaseModule = Join-Path $TempPath $Manifest.RootModule
        
        Write-Verbose "       Setting content for $ReleaseModule"
        Set-Content $ReleaseModule ((
            (Get-Content (Join-Path $ScriptRoot 'src\private\*.ps1') -Raw) + 
            (Get-Content (Join-Path $ScriptRoot 'src\public\*.ps1') -Raw)) -join "`r`n`r`n`r`n") -Encoding UTF8

        # just in case this was missing we update the manifest rootmodule.
        $Manifest.ModuleVersion = $Script:Version
        $Manifest.FunctionsToExport = $Script:FunctionsToExport

#        Update-Manifest $TempPath\$ModuleToBuild.psd1 -Property RootModule -Value $Manifest.RootModule

        # If there are any folders that aren't Public, Private, Tests, or Specs ...
        #$OtherFolders = Get-ChildItem $SourcePath -Directory -Exclude Public, Private, Tests, Specs
        # Then we need to copy everything in them
        #Copy-Item $OtherFolders -Recurse -Destination $TempPath

        # Finally, we need to copy any files in the Source directory
        #Get-ChildItem $ScriptRoot -File | 
        #    Where Name -ne $RootModule | 
        #    Copy-Item -Destination $TempPath

    }
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
        Get-ChildItem -Path $TempPath -Include "*.ps1","*.psm1" -Recurse -File | ForEach {
            $FormattedOutFile = $_.FullName
            Write-Output "Formatting File: $($FormattedOutFile)"
            $FormattedCode = get-content $_ -raw |
                Format-ScriptRemoveStatementSeparators |
                Format-ScriptExpandFunctionBlocks |
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

# Synopsis: Update module manifest with exported functions and version.
task UpdateModuleManifest -if ($Script:Manifest -ne $null){
    $Manifest.ModuleVersion = $Script:Version
    $Manifest.FunctionsToExport = $Script:FunctionsToExport
    $PrivateData = $null
    if ($Manifest.Keys -contains 'PrivateData') {
        $PrivateData = New-Object -TypeName psobject -Property $Manifest.PrivateData.PSdata
        $null = $Manifest.Remove('PrivateData')
    }

    New-ModuleManifest @Manifest -Path $TempPath\$ModuleToBuild.psd1 -PrivateData $PrivateData
}

# Synopsis: Create new release version directory from our temporary build directory and copy our results
task PublishModuleVersionRelease -If ( -not (Test-Path "$($ReleasePath)\$($Script:Version)")) {
    New-Item "$($ReleasePath)\$($Script:Version)" -ItemType:Directory
    Copy-Item $TempPath -Destination "$($ReleasePath)\$($Script:Version)" -Recurse
}

# Synopsis: Build the html help files and external module help files with PlatyPS
task CreateHelp {
	Import-Module PlatyPS
    
    # We need to import the module to give the helps functions something to work against
    Import-Module $ModuleFullPath

    # Create new html help files
    New-MarkdownHelp -module FormatPowerShellCode -OutputFolder "$($TempPath)\docs\" -Force
    
    # Create external help files
    New-ExternalHelp "$($TempPath)\docs" -OutputPath "$($TempPath)\en-US\"

    # Clean up loaded modules
    Remove-Module $ModuleToBuild
    Remove-Module PlatyPS
}

# Synopsis: Create a new version release directory for our release and copy our contents to it
task PushVersionRelease {
    $ThisReleasePath = "$ScriptRoot\release\$($Version)"
    Remove-Item $ThisReleasePath -Force -Recurse -ErrorAction 0
    New-Item $ThisReleasePath -ItemType:Directory
    Copy-Item -Path "$($TempPath)\*" -Destination $ThisReleasePath -Recurse
    Out-Zip $TempPath $ReleasePath\$ModuleToBuild'-'$Version'.zip' -overwrite
}

# Synopsis: Create the current release directory and copy this build to it.
task PushCurrentRelease {
    Remove-Item $CurrentReleasePath -Force -Recurse -ErrorAction 0
    New-Item $CurrentReleasePath -ItemType:Directory
    Copy-Item -Path "$($TempPath)\*" -Destination $CurrentReleasePath -Recurse
    Out-Zip $TempPath $ReleasePath\$ModuleToBuild'-current.zip' -overwrite
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
task . Version, Clean, Stage, CreateHelp,  FormatCode, UpdateModuleManifest, PushVersionRelease, PushCurrentRelease, GitStatus

# Synopsis: Test the code formatting module only
task TestModule Version, Clean, Stage, CreateHelp, FormatCode, UpdateModuleManifest

# Synopsis: Test combining the module into the temp build directory
task TestCombine Version, Clean, StageCombine, UpdateModuleManifest

# This psm1 file is purely for development. The build script will recreate this file entirely.

# Private and other methods and variables
Get-ChildItem "$PSScriptRoot\src\private","$PSScriptRoot\src\other" -Recurse -Filter "*.ps1" -File | Sort-Object Name | Foreach { 
    Write-Verbose "Dot sourcing private script file: $($_.Name)"
    . $_.FullName
}

# Load and export public methods
Get-ChildItem "$PSScriptRoot\src\public" -Recurse -Filter "*.ps1" -File | Sort-Object Name | Foreach { 
    Write-Verbose "Dot sourcing public script file: $($_.Name)"
    . $_.FullName

    # Find all the functions defined no deeper than the first level deep and export it.
    # This looks ugly but allows us to not keep any uneeded variables in memory that are not related to the module.
    ([System.Management.Automation.Language.Parser]::ParseInput((Get-Content -Path $_.FullName -Raw), [ref]$null, [ref]$null)).FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false) | Foreach {
        Export-ModuleMember $_.Name
    }
}
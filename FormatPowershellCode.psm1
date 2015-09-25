#region Private Variables
# Current script path
[string]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent
#endregion Private Variables

#region Methods
Get-ChildItem $ScriptPath/src/private -Recurse -Filter "*.ps1" -File | Foreach { 
    Write-Output "Dot sourcing private function: $($_.Name)"
    . $_.FullName
}

# Load and export methods
Get-ChildItem $ScriptPath/src/public -Recurse -Filter "*.ps1" -File | Foreach { 
    Write-Output "Dot sourcing public function: $($_.Name)"
    . $_.FullName
    Export-ModuleMember ($_.Name -replace '.ps1','')
}
#endregion Methods

#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    # cleanup when unloading module (if any)
}
#endregion Module Cleanup
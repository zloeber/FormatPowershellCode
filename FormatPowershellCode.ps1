#region Private Variables
# Current script path
[string]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent
#endregion Private Variables

#region Dependancies
Get-ChildItem $ScriptPath/src -Recurse -Filter "*.ps1" -File | Foreach { 
    Write-Verbose "Dot sourcing file: $($_.Name)"
    . $_.FullName
}
#endregion Depenencies

#region Module Export
Export-ModuleMember Format-Script*

#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    # Nothing
}
#endregion Module Cleanup
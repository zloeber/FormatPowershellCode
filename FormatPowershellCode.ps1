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

Export-ModuleMember Format-ScriptCondenseEnclosures
Export-ModuleMember Format-ScriptConvertKeywordsAndOperatorsToLower
Export-ModuleMember Format-ScriptExpandAliases
Export-ModuleMember Format-ScriptExpandTypeAccelerators
Export-ModuleMember Format-ScriptFormatCodeIndentation
Export-ModuleMember Format-ScriptFormatCommandNames
Export-ModuleMember Format-ScriptFormatTypeNames
Export-ModuleMember Format-ScriptGetKindLines
Export-ModuleMember Format-ScriptPadOperators
Export-ModuleMember Format-ScriptReduceLineLength
Export-ModuleMember Format-ScriptRemoveStatementSeparators
Export-ModuleMember Format-ScriptRemoveSuperfluousSpaces
Export-ModuleMember Format-ScriptReplaceHereStrings

#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    # Nothing
}
#endregion Module Cleanup
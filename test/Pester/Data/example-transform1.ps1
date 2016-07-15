$a = get-content 'C:\Users\rasputin\Dropbox\Zach_Docs\Projects\Git\ComputerAssetReport\New-AssetReportVersion2.ps1' -raw
$a = get-content 'C:\Users\rasputin\Dropbox\Zach_Docs\Projects\Git\FormatPowershellCode\tests\testcase-verylargefunction.ps1' -raw
$a | Format-ScriptFormatCommandNames -ExpandAliases -verbose | clip

#$a | Format-ScriptReplaceHereStrings | 
#    Format-ScriptExpandParameterBlocks | 
#    Format-ScriptExpandStatementBlocks | 
#    Format-ScriptExpandTypeAccelerators | 
#    Format-ScriptCondenseEnclosures | 
#    Format-ScriptFormatCommandNames -ExpandAliases | 
#    Format-ScriptExpandFunctionBlocks | 
#    Format-ScriptFormatTypeNames | 
#    Format-ScriptPadExpressions | 
#    Format-ScriptPadOperators | 
#    Format-ScriptRemoveStatementSeparators | 
#    Format-ScriptRemoveSuperfluousSpaces |
#    Format-ScriptReplaceHereStrings | 
#    Format-ScriptReduceLineLength |
#    Format-ScriptFormatCodeIndentation |
#    Format-ScriptReduceLineLength |
#    Format-ScriptFormatCodeIndentation |
#    clip
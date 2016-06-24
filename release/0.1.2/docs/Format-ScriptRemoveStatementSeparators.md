---
external help file: FormatPowerShellCode-help.xml
schema: 2.0.0
---

# Format-ScriptRemoveStatementSeparators
## SYNOPSIS
Finds all statement separators \(semicolons\) not in for loops and converts them to newlines.

## SYNTAX

```
Format-ScriptRemoveStatementSeparators [-Code] <String[]> [-SkipPostProcessingValidityCheck]
```

## DESCRIPTION
Finds all statement separators \(semicolons\) not in for loops and converts them to newlines.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$testfile = 'C:\temp\test.ps1'
```

PS \> $test = Get-Content $testfile -raw
PS \> $test | Format-RemoveStatementSeparators | clip

Description
-----------
Takes C:\temp\test.ps1 as input, remove statement separators and puts the result in the clipboard 
to be pasted elsewhere for review.

## PARAMETERS

### -Code
Multi-line or piped lines of code to process.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: 
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -SkipPostProcessingValidityCheck
After modifications have been made a check will be performed that the code has no errors.
Use this switch to bypass this check 
\(This is not recommended!\)

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0

Version History
1.0.0 - Initial release

## RELATED LINKS

[Online Version:]()


